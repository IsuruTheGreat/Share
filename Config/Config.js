/**
 * Config.js – central configuration (browser side, social-media only)
 * Keep secrets OUT of this file. Everything here is visible to anyone who opens the site.
 */
const CONFIG = {
  SCRIPT_URL: "https://script.google.com/macros/s/AKfycbxVgXF-ADOdfPelhSe8BWXpRAXd3xjuVqrc8mhv73ctuHhyeEMVUycqwi_7cHxvnCvaPg/exec",

  APP: {
    NAME: "Time Slot Scheduler",
    MAX_IMAGE_SIZE_MB: 10,
    MAX_VIDEO_SIZE_MB: 100,
    EDIT_IMAGE_MAX_SIZE_MB: 2,
    ALBUM_MAX_IMAGES: 10,         // users: Buffer's limit (Instagram, Facebook, TikTok)
    ALBUM_ADMIN_MAX_IMAGES: 10    // admins: same limit, Buffer does not accept more
  },

  // Club-level posts (BOD Post) and slot rules
  BOD_PROJECT: "BOD Post",
  SLOT_GRACE_MINUTES: 15,          // a slot closes 15 minutes after its start time

  // Fallback role lists (the owner edits the real lists in Owner Tools)
  ROLES: {
    PROJECT: ["Director", "Co-Director", "Chair", "Co-Chair", "Secretary", "Designer", "Caption Writer", "OC Member"],
    CLUB: ["President", "Vice President", "Treasurer", "Chief Media Officer", "Chief PR Officer", "PR Panel"]
  },
  SLOT_GROUPS: {
    "Morning": ["08:00 AM", "09:00 AM", "10:00 AM"],
    "Afternoon": ["12:00 PM", "01:00 PM", "02:00 PM"],
    "Evening": ["06:00 PM", "07:00 PM", "08:00 PM"]
  },
  SLOT_VALUES: {
    "08:00 AM": "8:00 AM", "09:00 AM": "9:00 AM", "10:00 AM": "10:00 AM",
    "12:00 PM": "12:00 PM", "01:00 PM": "1:00 PM", "02:00 PM": "2:00 PM",
    "06:00 PM": "6:00 PM", "07:00 PM": "7:00 PM", "08:00 PM": "8:00 PM"
  },
  FEATURES: {
    ENABLE_BUFFER_PUBLISH: true,
    ENABLE_MEDIA_DOWNLOAD: true
  },
  UI: {
    TOAST_DURATION_MS: 3000,
    SLOT_NO_LEADING_ZERO: false
  },
  LIMITS: {                       // text limits (Buffer)
    TIKTOK_PHOTO_TEXT: 4000,
    TIKTOK_VIDEO_TEXT: 2200,
    TIKTOK_TITLE: 90,
    YOUTUBE_TITLE: 100,
    YOUTUBE_DESCRIPTION: 5000
  },
  TITLE_SUFFIX: " | LeosofUOP"    // YouTube / TikTok title: Project name | LeosofUOP
};
Object.freeze(CONFIG.APP);
Object.freeze(CONFIG.ROLES.PROJECT);
Object.freeze(CONFIG.ROLES.CLUB);
Object.freeze(CONFIG.ROLES);
Object.freeze(CONFIG.SLOT_GROUPS);
Object.freeze(CONFIG.SLOT_VALUES);
Object.freeze(CONFIG.FEATURES);
Object.freeze(CONFIG.LIMITS);
Object.freeze(CONFIG);           // must be last