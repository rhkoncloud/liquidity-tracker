/* ---- Your settings. The ONLY value here is your public Microsoft Client ID. ----
   No passwords or secrets go here. Your OneDrive file path is intentionally NOT stored
   in this file — you enter it once inside the app (Settings ⚙) and it's saved locally
   on each device. That way this public repository never reveals where your data lives.

   clientId    : Application (client) ID from your Azure app registration (SETUP_GUIDE step 1).
                 This is a public identifier, not a secret — safe to commit.
   redirectUri : leave null to use the page's own URL (recommended). */
window.APP_CONFIG = {
  clientId: "de232d73-c2e1-4ac8-a0bc-60a95506de2a",
  redirectUri: null
};
