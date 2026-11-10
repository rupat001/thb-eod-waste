// ============================================================
//  SETTINGS — the only place the Sheet link and dates live.
//  Both the staff form (index.html) and manager page read this.
// ============================================================
window.EOD_CONFIG = {
  // Google Apps Script web app link (Deploy > Manage deployments). Ends in /exec.
  SCRIPT_URL: "https://script.google.com/macros/s/AKfycbxLboExVZ5exjDU9P-QLIDOm4zKs_bo-S_Asiz37Pd8Z_3d53rpcnyGm1q1jnB482BbwQ/exec",

  // First day entries are allowed / checked for missing days.
  START_DATE: "2026-09-01",

  // The pages and Code.gs agree on this. Bump only for a breaking change.
  API_VERSION: 2
};
