# Excel Data Cleaner - Office.js Add-in

A professional Excel Office.js Add-in that cleans and standardizes messy spreadsheet data using deterministic logic, with optional AI-powered analysis capabilities.

## Overview

Excel Data Cleaner is designed for non-technical Excel users who work with messy datasets. It provides a simple, one-click solution to clean and standardize data, making it ready for analysis or reporting.

### Purpose

- **Clean messy data**: Remove duplicates, trim whitespace, normalize formatting
- **Standardize format**: Consistent casing, empty cell handling, header detection
- **Professional output**: Auto-formatted headers, properly sized columns
- **Optional AI insights**: Get AI-powered analysis of your data (requires API configuration)

## Features

### Core Cleaning Features (Non-AI)

- ✅ **Trim whitespace** from all cells
- ✅ **Normalize text casing** to Title Case
- ✅ **Remove duplicate rows** (exact matching)
- ✅ **Remove fully empty rows**
- ✅ **Replace empty cells** with "N/A"
- ✅ **Detect and format header row** (bold + background color)
- ✅ **Auto-fit columns** after cleaning
- ✅ **Validate selection** (error if no range selected)

### Optional AI Analysis (Advanced)

- 🔄 **Toggle switch** to enable/disable AI insights
- 📊 **Sample analysis** of cleaned data (first 20 rows)
- 🤖 **AI-powered insights** (inconsistencies, anomalies, suggestions)
- 🔒 **Secure API integration** (API key via environment/config)
- ⚠️ **Graceful degradation** if AI is unavailable

## Tech Stack

- **Language**: JavaScript (ES6+)
- **Platform**: Office.js (Excel JavaScript API)
- **UI**: HTML + CSS (vanilla, no frameworks)
- **Async handling**: async/await
- **Optional AI**: OpenAI-compatible REST API (abstracted)
- **Tooling**: npm, Node.js
- **Version control**: Git-ready

## Project Structure

```
excel-data-cleaner-officejs/
├── src/
│   ├── taskpane/
│   │   ├── taskpane.html      # Main UI structure
│   │   ├── taskpane.css       # Styling
│   │   └── taskpane.js        # UI logic and orchestration
│   ├── commands/
│   │   └── commands.js        # Ribbon command handlers
│   └── utils/
│       ├── excelUtils.js      # Excel API wrapper functions
│       ├── dataCleaner.js     # Data transformation logic
│       └── aiAnalyzer.js      # AI integration (optional)
├── manifest.xml               # Office.js add-in manifest
├── package.json               # npm configuration
├── README.md                  # This file
└── .gitignore                 # Git exclusions
```

### File Descriptions

- **manifest.xml**: Office.js add-in configuration and metadata
- **taskpane.html/css/js**: Main user interface
- **excelUtils.js**: Wrapper functions for Excel JavaScript API operations
- **dataCleaner.js**: Pure functions for data cleaning (no Excel dependencies)
- **aiAnalyzer.js**: Optional AI analysis module with API abstraction
- **commands.js**: Ribbon button command handlers

## Getting Started

### Prerequisites

- **Node.js** (v14.0.0 or higher)
- **Excel** (Office 365, Excel 2016 or later)
- **npm** (comes with Node.js)

### Installation

1. **Clone or download this repository**
   ```bash
   cd excel-data-cleaner-officejs
   ```

2. **Install dependencies**
   ```bash
   npm install
   ```

3. **Install Office Add-in development certificates**
   ```bash
   npm run dev
   ```
   This will install SSL certificates needed for local development.

### Running Locally

1. **Start the development server**
   ```bash
   npm start
   ```
   Or:
   ```bash
   npm run dev
   ```

2. **Sideload the add-in in Excel**

   **Option A: Excel on Windows/Mac**
   - Open Excel
   - Go to **Insert** > **Add-ins** > **My Add-ins**
   - Click **Upload My Add-in**
   - Browse to the `manifest.xml` file in this project
   - Click **Upload**

   **Option B: Excel Online**
   - Open Excel Online
   - Go to **Insert** > **Add-ins** > **Upload My Add-in**
   - Upload the `manifest.xml` file

   **Option C: Using Office Add-in CLI (if installed)**
   ```bash
   office-addin-serve start manifest.xml
   ```

3. **Use the add-in**
   - Select a range of data in Excel
   - Open the "Excel Data Cleaner" task pane (via ribbon button or Insert > Add-ins)
   - Click "Clean Selected Data"
   - Review the cleaned results

### Development Workflow

1. Make changes to source files
2. The development server will automatically reload
3. Refresh the Excel add-in task pane to see changes

## AI Integration (Optional)

The add-in includes optional AI-powered analysis. This feature is **disabled by default** and requires configuration.

### Setting Up AI Analysis

1. **Get an API key** from an OpenAI-compatible API provider
   - OpenAI: https://platform.openai.com/api-keys
   - Or any OpenAI-compatible API endpoint

2. **Configure the API key**

   **Option A: Environment Variable (Recommended)**
   ```bash
   export AI_API_KEY="your-api-key-here"
   export AI_API_ENDPOINT="https://api.openai.com/v1/chat/completions"
   export AI_MODEL="gpt-3.5-turbo"
   ```

   **Option B: Configuration File**
   - Create a `config.json` file (not included in repo)
   - Add your API key (ensure it's in `.gitignore`)
   - Modify `aiAnalyzer.js` to read from the config file

3. **Enable AI in the UI**
   - Check the "Enable AI Insights (Optional)" checkbox
   - Click "Clean Selected Data"
   - AI insights will appear in the success message (if available)

### AI Analysis Details

- **Sample size**: First 20 rows of cleaned data
- **API endpoint**: Configurable (default: OpenAI)
- **Model**: Configurable (default: gpt-3.5-turbo)
- **Graceful failure**: Add-in works normally if AI is unavailable

### Security Notes

- ⚠️ **Never commit API keys** to version control
- ⚠️ **Use environment variables** or secure config files
- ⚠️ **For production**: Implement proper key management (Azure Key Vault, etc.)

## Usage

1. **Select data** in Excel (any range)
2. **Open the task pane** via the ribbon button
3. **Optionally enable AI insights** (requires API key)
4. **Click "Clean Selected Data"**
5. **Review results** in the status message

### What Gets Cleaned

- Leading/trailing whitespace removed
- Text converted to Title Case
- Duplicate rows removed
- Empty rows removed
- Empty cells replaced with "N/A"
- Header row formatted (bold, colored background)
- Columns auto-fitted to content

## Screenshots

_Placeholder for screenshots of the add-in in action_

- Task pane UI
- Before/after data comparison
- AI insights example

## Troubleshooting

### "No range selected" error
- Make sure you've selected a range of cells in Excel before clicking "Clean Selected Data"

### Add-in doesn't load
- Check that the development server is running
- Verify the manifest.xml is valid: `npm run validate`
- Check browser console for errors

### AI analysis not working
- Verify API key is configured correctly
- Check network connectivity
- Review browser console for API errors
- AI analysis is optional - core cleaning still works

### Certificate errors
- Run `npm run dev` to install development certificates
- On Windows, you may need to run as administrator

## Contributing

Contributions are welcome! Please ensure:
- Code follows the existing style
- Functions are well-documented
- Error handling is robust
- No hardcoded secrets or API keys

## License

MIT License - see LICENSE file for details

## Disclaimer

**AI Integration**: The optional AI analysis feature requires an external API and API key. This feature is provided as-is and users are responsible for:
- Securing their API keys
- Understanding API usage costs
- Complying with their API provider's terms of service
- Data privacy considerations when sending data to external APIs

The core data cleaning functionality works independently and does not require any external services.

## Support

For issues, questions, or contributions:
- Open an issue on GitHub
- Review the code comments for implementation details
- Check Office.js documentation: https://docs.microsoft.com/en-us/office/dev/add-ins/

---

**Built with Office.js and modern JavaScript** | **Production-ready** | **Interview-ready code quality**
