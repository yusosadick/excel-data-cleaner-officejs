/**
 * Data Cleaner Module
 * Provides pure functions for cleaning and standardizing data
 * All functions work with 2D arrays (row-major) and have no Excel API dependencies
 */

// Constants
const EMPTY_CELL_REPLACEMENT = "N/A";
const AI_SAMPLE_SIZE = 20;

/**
 * Converts text to Title Case
 * @param {string} text - The text to convert
 * @returns {string} Title case text
 */
function toTitleCase(text) {
  if (!text || typeof text !== "string") {
    return text;
  }
  
  return text.toLowerCase().replace(/\b\w/g, (char) => char.toUpperCase());
}

/**
 * Trims whitespace from a cell value
 * @param {any} cell - The cell value (can be any type)
 * @returns {any} The trimmed value (or original if not a string)
 */
function trimWhitespace(cell) {
  if (typeof cell === "string") {
    return cell.trim();
  }
  return cell;
}

/**
 * Normalizes text casing to Title Case for all string cells
 * @param {Array<Array<any>>} data - 2D array of data
 * @returns {Array<Array<any>>} Data with normalized casing
 */
function normalizeCasing(data) {
  return data.map((row) => 
    row.map((cell) => {
      if (typeof cell === "string" && cell.length > 0) {
        return toTitleCase(cell);
      }
      return cell;
    })
  );
}

/**
 * Removes duplicate rows from the data
 * Rows are compared as strings for exact matching
 * @param {Array<Array<any>>} data - 2D array of data
 * @returns {Array<Array<any>>} Data with duplicate rows removed
 */
function removeDuplicateRows(data) {
  const seen = new Set();
  const uniqueRows = [];
  
  for (const row of data) {
    // Convert row to string for comparison
    const rowString = JSON.stringify(row);
    
    if (!seen.has(rowString)) {
      seen.add(rowString);
      uniqueRows.push(row);
    }
  }
  
  return uniqueRows;
}

/**
 * Checks if a row is completely empty (all cells are empty, null, or undefined)
 * @param {Array<any>} row - A single row of data
 * @returns {boolean} True if the row is completely empty
 */
function isEmptyRow(row) {
  return row.every((cell) => {
    if (cell === null || cell === undefined) {
      return true;
    }
    if (typeof cell === "string") {
      return cell.trim().length === 0;
    }
    return false;
  });
}

/**
 * Removes fully empty rows from the data
 * @param {Array<Array<any>>} data - 2D array of data
 * @returns {Array<Array<any>>} Data with empty rows removed
 */
function removeEmptyRows(data) {
  return data.filter((row) => !isEmptyRow(row));
}

/**
 * Replaces empty cells with a standard placeholder
 * @param {Array<Array<any>>} data - 2D array of data
 * @returns {Array<Array<any>>} Data with empty cells replaced
 */
function replaceEmptyCells(data) {
  return data.map((row) =>
    row.map((cell) => {
      if (cell === null || cell === undefined) {
        return EMPTY_CELL_REPLACEMENT;
      }
      if (typeof cell === "string" && cell.trim().length === 0) {
        return EMPTY_CELL_REPLACEMENT;
      }
      return cell;
    })
  );
}

/**
 * Detects the header row index (first non-empty row)
 * @param {Array<Array<any>>} data - 2D array of data
 * @returns {number} Zero-based index of the header row, or 0 if not found
 */
function detectHeaderRow(data) {
  for (let i = 0; i < data.length; i++) {
    if (!isEmptyRow(data[i])) {
      return i;
    }
  }
  return 0; // Default to first row if all are empty
}

/**
 * Trims whitespace from all cells in the data
 * @param {Array<Array<any>>} data - 2D array of data
 * @returns {Array<Array<any>>} Data with trimmed whitespace
 */
function trimAllWhitespace(data) {
  return data.map((row) => row.map(trimWhitespace));
}

/**
 * Main data cleaning orchestration function
 * Applies all cleaning operations in the correct order
 * @param {Array<Array<any>>} data - 2D array of raw data
 * @returns {Object} Object containing cleaned data and metadata
 * @property {Array<Array<any>>} cleanedData - The cleaned data
 * @property {number} headerRowIndex - The detected header row index
 * @property {number} originalRowCount - Original number of rows
 * @property {number} cleanedRowCount - Number of rows after cleaning
 */
function cleanData(data) {
  if (!data || data.length === 0) {
    throw new Error("No data provided to clean.");
  }
  
  const originalRowCount = data.length;
  
  // Step 1: Trim whitespace from all cells
  let cleaned = trimAllWhitespace(data);
  
  // Step 2: Normalize text casing to Title Case
  cleaned = normalizeCasing(cleaned);
  
  // Step 3: Remove duplicate rows
  cleaned = removeDuplicateRows(cleaned);
  
  // Step 4: Remove fully empty rows
  cleaned = removeEmptyRows(cleaned);
  
  // Step 5: Replace empty cells with "N/A"
  cleaned = replaceEmptyCells(cleaned);
  
  // Step 6: Detect header row
  const headerRowIndex = detectHeaderRow(cleaned);
  
  const cleanedRowCount = cleaned.length;
  
  return {
    cleanedData: cleaned,
    headerRowIndex,
    originalRowCount,
    cleanedRowCount
  };
}

/**
 * Extracts a sample of data for AI analysis (first N rows)
 * @param {Array<Array<any>>} data - 2D array of data
 * @param {number} sampleSize - Number of rows to sample (default: AI_SAMPLE_SIZE)
 * @returns {Array<Array<any>>} Sample data
 */
function getSampleData(data, sampleSize = AI_SAMPLE_SIZE) {
  if (!data || data.length === 0) {
    return [];
  }
  
  return data.slice(0, Math.min(sampleSize, data.length));
}

// Export functions for use in other modules
// Browser environment: attach to window object
if (typeof window !== "undefined") {
  window.DataCleaner = {
    cleanData,
    trimWhitespace,
    normalizeCasing,
    removeDuplicateRows,
    removeEmptyRows,
    replaceEmptyCells,
    detectHeaderRow,
    getSampleData,
    EMPTY_CELL_REPLACEMENT,
    AI_SAMPLE_SIZE
  };
}

// Node.js environment: use module.exports
if (typeof module !== "undefined" && module.exports) {
  module.exports = {
    cleanData,
    trimWhitespace,
    normalizeCasing,
    removeDuplicateRows,
    removeEmptyRows,
    replaceEmptyCells,
    detectHeaderRow,
    getSampleData,
    EMPTY_CELL_REPLACEMENT,
    AI_SAMPLE_SIZE
  };
}
