/**
 * Excel Data Cleaner - Task Pane Main Script
 * Orchestrates the data cleaning workflow
 */

/* global Office, Excel */

// Initialize Office.js when the add-in loads
Office.onReady((info) => {
  if (info.host === Office.HostType.Excel) {
    // Add-in is ready
    initializeUI();
  } else {
    showStatus("This add-in only works in Excel.", "error");
  }
});

/**
 * Initializes the UI event handlers
 */
function initializeUI() {
  const cleanButton = document.getElementById("cleanButton");
  const aiToggle = document.getElementById("aiToggle");
  
  if (cleanButton) {
    cleanButton.addEventListener("click", handleCleanData);
  }
  
  // AI toggle doesn't need an event handler - it's checked when cleaning
}

/**
 * Main handler for the "Clean Selected Data" button click
 * 
 * Office.js Workflow:
 * 1. Gets the selected range (requires user to select cells first)
 * 2. Reads data using Excel.run batch context
 * 3. Cleans data using pure JavaScript functions (no Excel API)
 * 4. Writes cleaned data back to Excel
 * 5. Applies formatting (header row styling, column auto-fit)
 * 6. Optionally calls AI API for insights (gracefully degrades if unavailable)
 */
async function handleCleanData() {
  const cleanButton = document.getElementById("cleanButton");
  const aiToggle = document.getElementById("aiToggle");
  const statusMessage = document.getElementById("statusMessage");
  
  // Disable button and show loading state
  cleanButton.disabled = true;
  cleanButton.classList.add("loading");
  hideStatus();
  
  try {
    // Step 1: Get selected range
    showStatus("Reading selected range...", "info");
    const range = await getSelectedRange();
    const rangeAddress = await getRangeAddress(range);
    
    // Step 2: Read data from Excel
    showStatus("Reading data from Excel...", "info");
    const rawData = await readRangeData(range);
    
    if (!rawData || rawData.length === 0) {
      throw new Error("Selected range contains no data.");
    }
    
    // Step 3: Clean the data
    showStatus("Cleaning data...", "info");
    const cleaningResult = cleanData(rawData);
    const { cleanedData, headerRowIndex, originalRowCount, cleanedRowCount } = cleaningResult;
    
    // Step 4: Write cleaned data back to Excel
    showStatus("Writing cleaned data to Excel...", "info");
    await writeRangeData(range, cleanedData);
    
    // Step 5: Format header row
    showStatus("Formatting header row...", "info");
    await formatHeaderRow(range, headerRowIndex);
    
    // Step 6: Auto-fit columns
    showStatus("Auto-fitting columns...", "info");
    await autoFitColumns(range);
    
    // Step 7: Optional AI analysis
    const aiEnabled = aiToggle && aiToggle.checked;
    let aiInsights = null;
    
    if (aiEnabled) {
      showStatus("Analyzing data with AI (this may take a moment)...", "info");
      try {
        aiInsights = await analyzeData(cleanedData, true);
      } catch (error) {
        console.warn("AI analysis failed:", error);
        // Continue without AI insights
      }
    }
    
    // Step 8: Show success message
    const rowDiff = originalRowCount - cleanedRowCount;
    let successMessage = `Data cleaned successfully!\n\n`;
    successMessage += `Range: ${rangeAddress}\n`;
    successMessage += `Original rows: ${originalRowCount}\n`;
    successMessage += `Cleaned rows: ${cleanedRowCount}`;
    
    if (rowDiff > 0) {
      successMessage += `\nRemoved: ${rowDiff} duplicate or empty row(s)`;
    }
    
    if (aiInsights) {
      successMessage += `\n\nAI Insights:\n${aiInsights}`;
    } else if (aiEnabled) {
      successMessage += `\n\nAI analysis was requested but is not available. Check API key configuration.`;
    }
    
    showStatus(successMessage, "success");
    
  } catch (error) {
    console.error("Error cleaning data:", error);
    showStatus(`Error: ${error.message}`, "error");
  } finally {
    // Re-enable button
    cleanButton.disabled = false;
    cleanButton.classList.remove("loading");
  }
}

/**
 * Shows a status message to the user
 * @param {string} message - The message to display
 * @param {string} type - Message type: "success", "error", or "info"
 */
function showStatus(message, type = "info") {
  const statusMessage = document.getElementById("statusMessage");
  if (!statusMessage) {
    return;
  }
  
  // Remove all type classes
  statusMessage.classList.remove("success", "error", "info");
  
  // Add the new type class
  statusMessage.classList.add(type, "show");
  
  // Set the message text
  statusMessage.textContent = message;
  
  // Scroll into view
  statusMessage.scrollIntoView({ behavior: "smooth", block: "nearest" });
}

/**
 * Hides the status message
 */
function hideStatus() {
  const statusMessage = document.getElementById("statusMessage");
  if (statusMessage) {
    statusMessage.classList.remove("show");
  }
}

// Import utility functions from global namespaces (set by script tags)
// Use safe destructuring with fallbacks
const ExcelUtils = window.ExcelUtils || {};
const DataCleaner = window.DataCleaner || {};
const AIAnalyzer = window.AIAnalyzer || {};

const getSelectedRange = ExcelUtils.getSelectedRange;
const readRangeData = ExcelUtils.readRangeData;
const writeRangeData = ExcelUtils.writeRangeData;
const formatHeaderRow = ExcelUtils.formatHeaderRow;
const autoFitColumns = ExcelUtils.autoFitColumns;
const getRangeAddress = ExcelUtils.getRangeAddress;

const cleanData = DataCleaner.cleanData;
const analyzeData = AIAnalyzer.analyzeData;

// Validate that required functions are available
if (!getSelectedRange || !readRangeData || !writeRangeData) {
  console.error("ExcelUtils not loaded. Check script tags in HTML.");
}

if (!cleanData) {
  console.error("DataCleaner not loaded. Check script tags in HTML.");
}
