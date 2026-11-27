/**
 * Upload.js – signed Cloudinary uploads, shared by the Submit, Admin and Diagnostics pages.
 *
 * - The browser asks the Apps Script for a one-time signature (logged-in users only),
 *   so the upload preset no longer has to be public.
 * - Uses the /auto/upload address (images and videos).
 * - Videos over 20 MB are sent in pieces; each piece is retried on its own.
 * - Network errors are retried automatically (3 tries). The page is never reloaded.
 *
 * Needs Config.js to be loaded first (CONFIG.SCRIPT_URL).
 *
 * Usage:
 *   const url  = await uploadMedia(file, { onProgress: pct => {}, signal: abortController.signal });
 *   const data = await uploadMedia(file, { full: true });      // whole Cloudinary answer
 *   uploadToCloudinary(file, onProgress)                         // old name, still works
 */
(function () {
  "use strict";

  const CHUNK_FROM = 20 * 1024 * 1024;   // videos larger than this are sent in pieces
  const CHUNK_SIZE = 6 * 1024 * 1024;    // Cloudinary needs 5 MB or more per piece (except the last)
  const MAX_TRIES  = 3;

  function fail(message, o) {
    const e = new Error(message);
    e.retry   = !!(o && o.retry);
    e.aborted = !!(o && o.aborted);
    return e;
  }
  const sleep = (ms) => new Promise(r => setTimeout(r, ms));
  const isVideo = (f) => /^video\//.test(f.type || "");

  function authToken() {
    try { return (JSON.parse(sessionStorage.getItem("auth")) || {}).token || ""; } catch (e) { return ""; }
  }

  // One-time upload permission from the server (only for logged-in users)
  async function getSignature(signal, file) {
    let res, d;
    try {
      res = await fetch(CONFIG.SCRIPT_URL, {
        method: "POST",
        headers: { "Content-Type": "text/plain;charset=utf-8" },
        body: JSON.stringify({ token: authToken(), action: "getUploadSignature", size: file ? file.size : 0, type: file ? file.type : "" }),
        signal: signal
      });
    } catch (e) {
      if (e && e.name === "AbortError") throw fail("Cancelled", { aborted: true });
      throw fail("Your internet is weak.", { retry: true });
    }
    try { d = await res.json(); } catch (e) { throw fail("Your internet is weak.", { retry: true }); }
    if (d && d.paused) throw fail(d.error || "The system is paused.");
    if (!d || !d.success) throw fail((d && d.error) || "Could not start the upload. Please log in again.");
    return d;
  }

  function buildForm(sig, blob, name) {
    const fd = new FormData();
    fd.append("api_key", sig.apiKey);
    fd.append("timestamp", sig.timestamp);
    fd.append("signature", sig.signature);
    if (sig.folder)       fd.append("folder", sig.folder);
    if (sig.publicId)     fd.append("public_id", sig.publicId);               // one file per permission
    if (sig.allowedFormats) fd.append("allowed_formats", sig.allowedFormats);   // part of the signature
    if (sig.uploadPreset) fd.append("upload_preset", sig.uploadPreset);
    fd.append("file", blob, name || "upload");
    return fd;
  }

  function send(url, form, headers, onBytes, signal) {
    return new Promise((resolve, reject) => {
      const xhr = new XMLHttpRequest();
      xhr.open("POST", url);
      Object.keys(headers || {}).forEach(k => xhr.setRequestHeader(k, headers[k]));
      xhr.timeout = 10 * 60 * 1000;
      if (xhr.upload && onBytes) xhr.upload.onprogress = (e) => { if (e.lengthComputable) onBytes(e.loaded, e.total); };
      xhr.onload = () => {
        let d = null;
        try { d = JSON.parse(xhr.responseText); } catch (e) {}
        if (xhr.status >= 200 && xhr.status < 300 && d) return resolve(d);
        const busy = xhr.status >= 500 || xhr.status === 429;
        reject(fail(busy ? "The upload server is busy." : ((d && d.error && d.error.message) || "The upload was refused."), { retry: busy }));
      };
      xhr.onerror   = () => reject(fail("Your internet is weak.", { retry: true }));
      xhr.ontimeout = () => reject(fail("Your internet is weak.", { retry: true }));
      xhr.onabort   = () => reject(fail("Cancelled", { aborted: true }));
      if (signal) {
        if (signal.aborted) return reject(fail("Cancelled", { aborted: true }));
        signal.addEventListener("abort", () => xhr.abort(), { once: true });
      }
      xhr.send(form);
    });
  }

  // Optional: shrink very large phone photos (JPEG only) before upload. Any problem = send the original.
  async function shrinkPhoto(file) {
    if (file.type !== "image/jpeg" || file.size < 3 * 1024 * 1024 || !window.createImageBitmap) return file;
    const bmp = await createImageBitmap(file);
    const k = Math.min(1, 2560 / Math.max(bmp.width, bmp.height));
    if (k === 1 && file.size < 6 * 1024 * 1024) { if (bmp.close) bmp.close(); return file; }
    const c = document.createElement("canvas");
    c.width = Math.round(bmp.width * k); c.height = Math.round(bmp.height * k);
    c.getContext("2d").drawImage(bmp, 0, 0, c.width, c.height);
    if (bmp.close) bmp.close();
    const blob = await new Promise(r => c.toBlob(r, "image/jpeg", 0.88));
    c.width = c.height = 0;
    if (!blob || blob.size >= file.size) return file;
    return new File([blob], file.name, { type: "image/jpeg" });
  }

  async function attempt(file, opt, setPct) {
    const sig = await getSignature(opt.signal, file);
    const url = "https://api.cloudinary.com/v1_1/" + sig.cloudName + "/auto/upload";

    if (!isVideo(file) || file.size <= CHUNK_FROM) {
      return send(url, buildForm(sig, file, file.name), null, (l, t) => setPct(l / t * 100), opt.signal);
    }

    // Big video: send in pieces, retry each piece on its own
    const id = "u" + Date.now() + Math.random().toString(36).slice(2);
    let start = 0, last = null;
    while (start < file.size) {
      const end = Math.min(start + CHUNK_SIZE, file.size);
      const piece = file.slice(start, end);
      const headers = { "X-Unique-Upload-Id": id, "Content-Range": "bytes " + start + "-" + (end - 1) + "/" + file.size };
      let tries = 0;
      for (;;) {
        try {
          last = await send(url, buildForm(sig, piece, file.name), headers, (l) => setPct((start + Math.min(l, end - start)) / file.size * 100), opt.signal);
          break;
        } catch (e) {
          if (e.aborted || !e.retry) throw e;
          if (++tries >= MAX_TRIES) { e.retry = false; throw e; }
          await sleep(1000 * tries);
        }
      }
      start = end;
    }
    return last;
  }

  async function uploadMedia(file, opt) {
    opt = opt || {};
    const onProgress = opt.onProgress || function () {};
    let src = file;
    try { src = await shrinkPhoto(file); } catch (e) { src = file; }

    let tries = 0;
    for (;;) {
      try {
        const d = await attempt(src, opt, (p) => onProgress(Math.max(0, Math.min(99, p))));
        if (!d || !d.secure_url) throw fail((d && d.error && d.error.message) || "The upload failed.");
        onProgress(100);
        return opt.full ? d : d.secure_url;
      } catch (e) {
        if (e.aborted || !e.retry || ++tries >= MAX_TRIES) throw e;
        if (opt.onRetry) opt.onRetry(tries);
        onProgress(0);
        await sleep(1500 * tries);
        if (opt.signal && opt.signal.aborted) throw fail("Cancelled", { aborted: true });
      }
    }
  }

  window.uploadMedia = uploadMedia;
  window.uploadToCloudinary = function (file, onProgress) { return uploadMedia(file, { onProgress: onProgress }); };   // old name
})();