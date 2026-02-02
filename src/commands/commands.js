/**
 * Office.js Commands Module
 * Handles ribbon command button actions
 */

/* global Office */

/**
 * Initializes the add-in when the ribbon loads
 */
Office.onReady((info) => {
  if (info.host === Office.HostType.Excel) {
    // Add-in is ready in Excel
  }
});

/**
 * Handler for the "Show Taskpane" button click
 * This function is called from the manifest.xml command definition
 */
function showTaskpane(event) {
  // The task pane is already shown via the manifest action
  // This function can be used for additional initialization if needed
  event.completed();
}

// Register the function with Office.js
if (typeof Office !== "undefined") {
  Office.actions.associate("showTaskpane", showTaskpane);
}
