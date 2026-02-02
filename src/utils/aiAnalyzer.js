/**
 * AI Analyzer Module
 * Provides optional AI-powered data analysis via external API
 * This module is fully optional and gracefully degrades if disabled or unavailable
 */

// Configuration constants
// In browser environment, these can be set via window.config or environment variables
// For production, use a proper configuration system
function getConfigValue(key, defaultValue) {
  // Check window.config first (for browser-based config)
  if (typeof window !== "undefined" && window.config && window.config[key]) {
    return window.config[key];
  }
  // Check process.env (for Node.js environments)
  if (typeof process !== "undefined" && process.env && process.env[key]) {
    return process.env[key];
  }
  return defaultValue;
}

const AI_API_ENDPOINT = getConfigValue("AI_API_ENDPOINT", "https://api.openai.com/v1/chat/completions");
const AI_MODEL = getConfigValue("AI_MODEL", "gpt-3.5-turbo");
const AI_SAMPLE_SIZE = 20; // Number of rows to send for analysis

/**
 * Retrieves the API key from environment variables or configuration
 * In a production environment, this should be securely stored and retrieved
 * @returns {string|null} The API key, or null if not configured
 */
function getApiKey() {
  // Check window.config first (for browser-based config)
  // WARNING: Storing API keys in client-side code is a security risk
  // For production, use a backend proxy or secure key management service
  if (typeof window !== "undefined" && window.config && window.config.AI_API_KEY) {
    return window.config.AI_API_KEY;
  }
  
  // Check environment variable (for Node.js/server environments)
  if (typeof process !== "undefined" && process.env && process.env.AI_API_KEY) {
    return process.env.AI_API_KEY;
  }
  
  // No API key found
  // For production: implement secure key management (Azure Key Vault, AWS Secrets Manager, etc.)
  return null;
}

/**
 * Formats data sample for AI analysis
 * Converts 2D array to a readable text format
 * @param {Array<Array<any>>} sampleData - Sample data rows
 * @returns {string} Formatted text representation
 */
function formatDataForAI(sampleData) {
  if (!sampleData || sampleData.length === 0) {
    return "No data available for analysis.";
  }
  
  // Convert to tab-separated format for readability
  const lines = sampleData.map((row) => row.join("\t"));
  return lines.join("\n");
}

/**
 * Calls the AI API to analyze the data sample
 * @param {Array<Array<any>>} sampleData - Sample data to analyze
 * @param {string} apiKey - API key for authentication
 * @returns {Promise<Object>} AI analysis response
 */
async function callAIApi(sampleData, apiKey) {
  const formattedData = formatDataForAI(sampleData);
  
  const prompt = `Analyze the following spreadsheet data sample and provide insights:
- Identify any data inconsistencies or anomalies
- Suggest improvements for data quality
- Note any patterns or issues

Data sample:
${formattedData}

Please provide concise, actionable insights.`;

  const requestBody = {
    model: AI_MODEL,
    messages: [
      {
        role: "system",
        content: "You are a data quality analyst. Provide concise, actionable insights about spreadsheet data."
      },
      {
        role: "user",
        content: prompt
      }
    ],
    max_tokens: 500,
    temperature: 0.3
  };

  try {
    const response = await fetch(AI_API_ENDPOINT, {
      method: "POST",
      headers: {
        "Content-Type": "application/json",
        "Authorization": `Bearer ${apiKey}`
      },
      body: JSON.stringify(requestBody)
    });

    if (!response.ok) {
      const errorText = await response.text();
      throw new Error(`AI API error: ${response.status} - ${errorText}`);
    }

    const data = await response.json();
    return data;
  } catch (error) {
    console.error("AI API call failed:", error);
    throw error;
  }
}

/**
 * Parses the AI API response to extract insights
 * @param {Object} apiResponse - The response from the AI API
 * @returns {string} Formatted insights text
 */
function parseAIResponse(apiResponse) {
  if (!apiResponse || !apiResponse.choices || apiResponse.choices.length === 0) {
    return "No insights available from AI analysis.";
  }

  const content = apiResponse.choices[0].message?.content;
  if (!content) {
    return "AI response format unexpected.";
  }

  return content.trim();
}

/**
 * Main AI analysis function
 * Analyzes a sample of cleaned data and returns insights
 * @param {Array<Array<any>>} cleanedData - The cleaned data to analyze
 * @param {boolean} enabled - Whether AI analysis is enabled
 * @returns {Promise<string|null>} AI insights as a string, or null if disabled/failed
 */
async function analyzeData(cleanedData, enabled = false) {
  // Early return if AI is disabled
  if (!enabled) {
    return null;
  }

  // Check for API key
  const apiKey = getApiKey();
  if (!apiKey) {
    console.warn("AI analysis requested but no API key configured.");
    return null;
  }

  // Validate data
  if (!cleanedData || cleanedData.length === 0) {
    console.warn("No data provided for AI analysis.");
    return null;
  }

  try {
    // Extract sample (first N rows)
    const sampleData = cleanedData.slice(0, AI_SAMPLE_SIZE);
    
    // Call AI API
    const apiResponse = await callAIApi(sampleData, apiKey);
    
    // Parse and return insights
    const insights = parseAIResponse(apiResponse);
    return insights;
  } catch (error) {
    // Gracefully handle errors - don't throw, just log and return null
    console.error("AI analysis failed:", error.message);
    return null;
  }
}

// Export functions for use in other modules
// Browser environment: attach to window object
if (typeof window !== "undefined") {
  window.AIAnalyzer = {
    analyzeData,
    getApiKey,
    callAIApi,
    parseAIResponse,
    AI_API_ENDPOINT,
    AI_MODEL,
    AI_SAMPLE_SIZE
  };
}

// Node.js environment: use module.exports
if (typeof module !== "undefined" && module.exports) {
  module.exports = {
    analyzeData,
    getApiKey,
    callAIApi,
    parseAIResponse,
    AI_API_ENDPOINT,
    AI_MODEL,
    AI_SAMPLE_SIZE
  };
}
