// ===================================================================
// Menu.gs — Custom spreadsheet menu for AWA Admin Portal
//
// Adds an "AWA Admin" menu to the Google Sheet toolbar when the
// spreadsheet is opened. Requires the web app deployment URL below.
// ===================================================================

// Replace with your Apps Script Web App deployment URL.
// Deploy → Manage deployments → copy the Web App URL.
const PORTAL_URL = 'YOUR_DEPLOYMENT_URL_HERE';

/**
 * Runs automatically when the spreadsheet is opened.
 * Adds the "AWA Admin" custom menu to the toolbar.
 */
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('AWA Admin')
    .addItem('Open Check-In Portal', 'openAdminPortal')
    .addToUi();
}

/**
 * Opens the Admin Portal in a new browser tab.
 * Called by the "Open Check-In Portal" menu item.
 */
function openAdminPortal() {
  const html = HtmlService
    .createHtmlOutput(
      '<script>window.open("' + PORTAL_URL + '"); google.script.host.close();</script>'
    )
    .setWidth(10).setHeight(10);
  SpreadsheetApp.getUi().showModalDialog(html, 'Opening Admin Portal…');
}
