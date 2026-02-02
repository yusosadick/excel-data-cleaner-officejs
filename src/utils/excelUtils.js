/**
 * Excel Utilities Module
 * Provides wrapper functions for Excel JavaScript API operations
 */

/**
 * Gets the currently selected range in Excel
 * 
 * Office.js Note: Excel.run creates a request context that batches all operations.
 * The returned Range object is a proxy that can be used in subsequent Excel.run calls.
 * 
 * @returns {Promise<Excel.Range>} The selected range (proxy object for use in Excel.run)
 * @throws {Error} If no range is selected or selection is invalid
 */
async function getSelectedRange() {
  return Excel.run(async (context) => {
    const range = context.workbook.getSelectedRange();
    range.load("address");
    await context.sync();
    
    if (!range.address) {
      throw new Error("No range selected. Please select a range of cells to clean.");
    }
    
    return range;
  });
}

/**
 * Reads data from the specified Excel range
 * 
 * Office.js Note: range.values returns a 2D array where each inner array represents a row.
 * Empty cells are returned as null. Numbers, dates, and formulas return their calculated values.
 * 
 * @param {Excel.Range} range - The Excel range to read from (proxy object from Excel.run)
 * @returns {Promise<Array<Array<any>>>} 2D array of cell values (row-major format)
 */
async function readRangeData(range) {
  return Excel.run(async (context) => {
    range.load("values");
    await context.sync();
    
    // Convert to plain JavaScript array
    // Note: range.values is already a plain JS array after context.sync()
    const values = range.values;
    return values;
  });
}

/**
 * Writes data to the specified Excel range
 * 
 * Office.js Note: Automatically resizes the range to match data dimensions.
 * The range proxy object can be used across different Excel.run contexts.
 * 
 * @param {Excel.Range} range - The Excel range to write to (proxy object from Excel.run)
 * @param {Array<Array<any>>} data - 2D array of values to write (row-major format)
 * @throws {Error} If data is empty or invalid
 */
async function writeRangeData(range, data) {
  return Excel.run(async (context) => {
    // Resize range if needed to match data dimensions
    const rowCount = data.length;
    const colCount = data.length > 0 ? data[0].length : 0;
    
    if (rowCount === 0 || colCount === 0) {
      throw new Error("Cannot write empty data to Excel range.");
    }
    
    // Get the starting cell of the range and resize to match data
    // Note: getResizedRange parameters are relative offsets, not absolute sizes
    const startCell = range.getCell(0, 0);
    const newRange = startCell.getResizedRange(rowCount - 1, colCount - 1);
    
    // Write the data (Office.js automatically handles type conversion)
    newRange.values = data;
    await context.sync();
  });
}

/**
 * Formats the header row (first row) with bold text and background color
 * @param {Excel.Range} range - The Excel range containing the data
 * @param {number} headerRowIndex - Zero-based index of the header row (default: 0)
 */
async function formatHeaderRow(range, headerRowIndex = 0) {
  return Excel.run(async (context) => {
    // Get the header row
    const headerRow = range.getRow(headerRowIndex);
    
    // Apply formatting
    headerRow.format.font.bold = true;
    headerRow.format.fill.color = "#4472C4"; // Professional blue background
    headerRow.format.font.color = "#FFFFFF"; // White text
    
    await context.sync();
  });
}

/**
 * Auto-fits columns in the specified range to fit their content
 * 
 * Office.js Note: autofitColumns() adjusts column width based on cell content.
 * This operation requires loading the columns collection first.
 * 
 * @param {Excel.Range} range - The Excel range to auto-fit (proxy object from Excel.run)
 */
async function autoFitColumns(range) {
  return Excel.run(async (context) => {
    // Get the columns in the range
    const columns = range.getColumns();
    columns.load("columnIndex");
    await context.sync();
    
    // Auto-fit each column
    // Note: autofitColumns() is called on each column individually
    columns.items.forEach((column) => {
      column.format.autofitColumns();
    });
    
    await context.sync();
  });
}

/**
 * Gets the address of a range as a string
 * @param {Excel.Range} range - The Excel range
 * @returns {Promise<string>} The address string (e.g., "A1:C10")
 */
async function getRangeAddress(range) {
  return Excel.run(async (context) => {
    range.load("address");
    await context.sync();
    return range.address;
  });
}

// Export functions for use in other modules
// Browser environment: attach to window object
if (typeof window !== "undefined") {
  window.ExcelUtils = {
    getSelectedRange,
    readRangeData,
    writeRangeData,
    formatHeaderRow,
    autoFitColumns,
    getRangeAddress
  };
}

// Node.js environment: use module.exports
if (typeof module !== "undefined" && module.exports) {
  module.exports = {
    getSelectedRange,
    readRangeData,
    writeRangeData,
    formatHeaderRow,
    autoFitColumns,
    getRangeAddress
  };
}
