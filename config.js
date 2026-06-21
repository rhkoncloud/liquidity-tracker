/* ---- Your settings. Edit the two values below. No passwords or secrets go here. ----
   clientId   : the Application (client) ID from your Azure app registration (see SETUP_GUIDE step 1).
   filePath   : path to your workbook inside OneDrive, relative to the OneDrive root.
   redirectUri: leave null to use the page's own URL (recommended). Must match what you
                register in Azure exactly. */
window.APP_CONFIG = {
  clientId: "PASTE_YOUR_CLIENT_ID_HERE",
  filePath: "Details/Finances/Finance_Review_FY2025-26.xlsx",
  redirectUri: null
};
