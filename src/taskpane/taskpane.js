/* global console, document, Excel, Office */

let currentResults = "";

Office.onReady((info) => {
  if (info.host === Office.HostType.Excel) {
    console.log("Excel Claude Assistant loaded successfully!");
    
    // Initialize event listeners
    document.getElementById("analyze-selection").onclick = () => analyzeSelection();
    document.getElementById("create-formula").onclick = () => createFormula();
    document.getElementById("data-insights").onclick = () => getDataInsights();
    document.getElementById("auto-categorize").onclick = () => autoCategorizeData();
    document.getElementById("clean-data").onclick = () => cleanAndStandardizeData();
    document.getElementById("add-calculations").onclick = () => addCalculations();
    document.getElementById("send-custom").onclick = () => sendCustomRequest();
    document.getElementById("insert-results").onclick = () => insertResults();
    
    // Load saved API key
    const savedKey = localStorage.getItem("claude-api-key");
    if (savedKey) {
      document.getElementById("api-key").value = savedKey;
    }
    
    // Save API key when changed
    document.getElementById("api-key").onchange = () => {
      const apiKey = document.getElementById("api-key").value;
      localStorage.setItem("claude-api-key", apiKey);
    };
  }
});

async function analyzeSelection() {
  try {
    showLoading(true);
    showStatus("Getting selected data...", "info");
    
    await Excel.run(async (context) => {
      const range = context.workbook.getSelectedRange();
      range.load("values, address, rowCount, columnCount");
      
      await context.sync();
      
      if (range.rowCount === 1 && range.columnCount === 1) {
        showStatus("Please select a larger data range", "error");
        return;
      }
      
      const dataText = convertRangeToText(range.values);
      const prompt = "Analyze this Excel data and suggest NEW COLUMNS or CALCULATIONS to add:\n\nData:\n" + dataText + "\n\nPlease suggest 2-3 new columns. For each suggestion, provide:\nCOLUMN: [Header Name]\nFORMULA: [Excel formula]\nEXPLANATION: [Why this is useful]";

      const response = await callClaudeAPI(prompt);
      
      // Apply simple analysis to spreadsheet
      await applySimpleAnalysis(context, range, response);
      
      displayResults(response);
      showStatus("Analysis applied to spreadsheet!", "success");
    });
  } catch (error) {
    showStatus("Error: " + error.message, "error");
    console.error("Analysis error:", error);
  } finally {
    showLoading(false);
  }
}

async function createFormula() {
  try {
    showLoading(true);
    
    const task = prompt("What calculation or formula do you need help with?");
    if (!task) return;
    
    await Excel.run(async (context) => {
      const range = context.workbook.getSelectedRange();
      range.load("values, address, rowCount, columnCount");
      
      await context.sync();
      
      const dataText = convertRangeToText(range.values);
      const prompt = "Create Excel formula for this task: " + task + "\n\nData:\n" + dataText + "\n\nProvide:\nFORMULA: [Excel formula starting with =]\nHEADER: [Column header name]";

      const response = await callClaudeAPI(prompt);
      
      // Apply formula to spreadsheet
      await applyFormulaToSpreadsheet(context, range, response);
      
      displayResults(response);
      showStatus("Formula applied to spreadsheet!", "success");
    });
  } catch (error) {
    showStatus("Error: " + error.message, "error");
  } finally {
    showLoading(false);
  }
}

async function getDataInsights() {
  try {
    showLoading(true);
    
    await Excel.run(async (context) => {
      const range = context.workbook.getSelectedRange();
      range.load("values, address, rowCount, columnCount");
      
      await context.sync();
      
      const dataText = convertRangeToText(range.values);
      const prompt = "Analyze this data and create calculated columns:\n\n" + dataText + "\n\nCreate 2-3 columns with:\nCOLUMN: [Column Name]\nFORMULA: [Excel formula]\nEXPLANATION: [Business insight]";

      const response = await callClaudeAPI(prompt);
      
      // Apply insights as new columns
      await applySimpleAnalysis(context, range, response);
      
      displayResults(response);
      showStatus("Insights added as new columns!", "success");
    });
  } catch (error) {
    showStatus("Error: " + error.message, "error");
  } finally {
    showLoading(false);
  }
}

async function sendCustomRequest() {
  try {
    const customPrompt = document.getElementById("custom-prompt").value.trim();
    if (!customPrompt) {
      showStatus("Please enter a request", "error");
      return;
    }
    
    showLoading(true);
    
    const response = await callClaudeAPI(customPrompt);
    
    await Excel.run(async (context) => {
      // Create simple structured worksheet
      await createSimpleWorksheet(context, response, customPrompt);
      
      displayResults(response);
      showStatus("Request completed and applied to Excel!", "success");
    });
  } catch (error) {
    showStatus("Error: " + error.message, "error");
  } finally {
    showLoading(false);
  }
}

async function createSimpleWorksheet(context, response, title) {
  try {
    // Create new worksheet
    const timestamp = new Date().toISOString().slice(0, 16).replace(/[:-]/g, "");
    const worksheetName = "Claude_" + timestamp;
    
    const newWorksheet = context.workbook.worksheets.add(worksheetName);
    
    // Extract ONLY tables from code blocks - ignore all other text
    const tables = extractOnlyCodeBlockTables(response);
    
    if (tables.length === 0) {
      showStatus("No code block tables found in response", "info");
      return;
    }
    
    let currentRow = 0;
    
    // Insert only the extracted tables
    for (let table of tables) {
      currentRow = await insertCleanTable(context, newWorksheet, table, currentRow);
      currentRow += 2; // Add spacing between tables
    }
    
    // Basic formatting
    newWorksheet.getRange("A:A").format.columnWidth = 200;
    newWorksheet.getRange("B:F").format.columnWidth = 100;
    
    newWorksheet.activate();
    await context.sync();
    
  } catch (error) {
    console.error("Error creating worksheet:", error);
    showStatus("Error creating worksheet", "error");
  }
}

function extractOnlyCodeBlockTables(response) {
  const tables = [];
  const lines = response.split('\n');
  
  let inCodeBlock = false;
  let currentTable = [];
  let tableTitle = "";
  
  for (let i = 0; i < lines.length; i++) {
    const line = lines[i].trim();
    
    if (line === '```') {
      if (!inCodeBlock) {
        // Starting code block - look for title in previous lines
        inCodeBlock = true;
        currentTable = [];
        tableTitle = findTableTitle(lines, i);
      } else {
        // Ending code block - save table if it has data
        if (currentTable.length > 0) {
          tables.push({
            title: tableTitle || "Data",
            rows: currentTable
          });
        }
        inCodeBlock = false;
        currentTable = [];
        tableTitle = "";
      }
    } else if (inCodeBlock && line) {
      // Only process lines inside code blocks
      const cells = parseCleanRow(line);
      if (cells.length > 0) {
        currentTable.push(cells);
      }
    }
    // Completely ignore all lines outside code blocks
  }
  
  return tables;
}

function findTableTitle(lines, codeBlockIndex) {
  // Look backwards for section title (like "1. Revenue Model:")
  for (let i = codeBlockIndex - 1; i >= Math.max(0, codeBlockIndex - 3); i--) {
    const line = lines[i].trim();
    if (line) {
      // Remove numbering and colon
      let title = line.replace(/^\d+\.\s*/, '').replace(':', '').trim();
      if (title.length > 0 && title !== '```') {
        return title;
      }
    }
  }
  return "Table";
}

function parseCleanRow(line) {
  // Skip separator lines completely
  if (line.match(/^[-\s|=_]+$/)) {
    return [];
  }
  
  // Split by multiple spaces (2 or more) or tabs
  let cells = line.split(/\s{2,}|\t/).map(cell => cell.trim()).filter(cell => cell);
  
  // If no clear separation, try to parse structured data
  if (cells.length <= 1 && line.includes(' ')) {
    const parts = line.split(/\s+/);
    if (parts.length >= 2) {
      // Check if we have a label followed by numbers
      const hasNumbers = parts.slice(1).some(part => part.match(/[\d,%]/));
      if (hasNumbers) {
        // First part is label, rest might be data
        const label = parts[0];
        const data = parts.slice(1);
        cells = [label].concat(data);
      }
    }
  }
  
  // Clean up cells and handle Excel issues
  return cells.map(cell => cleanCellValue(cell));
}

function cleanCellValue(value) {
  if (!value || value === '') return '';
  
  // Handle the "-" issue - Excel treats "-" as formula
  if (value === '-') {
    return '—'; // Use em dash instead
  }
  
  // Handle cells that start with "-" but aren't formulas
  if (value.startsWith('-') && !value.match(/^-[\d,\.]+/)) {
    return '—' + value.slice(1); // Replace leading hyphen with em dash
  }
  
  // Clean up other problematic characters
  return value.trim();
}

async function insertCleanTable(context, worksheet, table, startRow) {
  let currentRow = startRow;
  
  // Insert title if it exists and is meaningful
  if (table.title && table.title !== "Data" && table.title !== "Table") {
    const titleCell = worksheet.getCell(currentRow, 0);
    titleCell.values = [[table.title]];
    titleCell.format.font.bold = true;
    titleCell.format.font.size = 12;
    titleCell.format.fill.color = "#4472C4";
    titleCell.format.font.color = "white";
    currentRow++;
  }
  
  // Insert rows
  for (let rowIndex = 0; rowIndex < table.rows.length; rowIndex++) {
    const row = table.rows[rowIndex];
    
    for (let colIndex = 0; colIndex < row.length; colIndex++) {
      const cell = worksheet.getCell(currentRow, colIndex);
      const cellValue = row[colIndex];
      
      // Handle numeric values
      const cleanNumValue = cellValue.replace(/,/g, '').replace(/[()]/g, '');
      const numValue = parseFloat(cleanNumValue);
      
      if (!isNaN(numValue) && cellValue.match(/[\d,]/)) {
        // It's a number
        cell.values = [[numValue]];
        
        // Format large numbers with commas
        if (Math.abs(numValue) >= 100) {
          cell.numberFormat = [["#,##0"]];
        }
        
        // Handle percentages
        if (cellValue.includes('%')) {
          cell.numberFormat = [["0%"]];
          cell.values = [[numValue / 100]];
        }
        
        // Right align numbers
        cell.format.horizontalAlignment = "Right";
      } 
      // Handle text values (including cleaned "-" characters)
      else {
        cell.values = [[cellValue]];
        
        // Left align text
        if (colIndex === 0) {
          cell.format.horizontalAlignment = "Left";
        }
      }
      
      // Format headers (first row or rows containing "Year")
      if (rowIndex === 0 || cellValue.includes('Year') || cellValue.includes('$')) {
        cell.format.font.bold = true;
        cell.format.fill.color = "#D9E2F3";
      }
    }
    
    currentRow++;
  }
  
  return currentRow;
}

function parseSimpleRow(line) {
  // This function is now replaced by parseCleanRow
  return parseCleanRow(line);
}

// Simplified functions for applying results
async function applySimpleAnalysis(context, originalRange, response) {
  try {
    const suggestions = parseColumnSuggestions(response);
    
    if (suggestions.length === 0) {
      showStatus("No actionable columns found in response", "error");
      return;
    }
    
    const startCol = originalRange.columnIndex + originalRange.columnCount;
    
    for (let i = 0; i < Math.min(suggestions.length, 3); i++) {
      const suggestion = suggestions[i];
      const colIndex = startCol + i;
      
      // Add header
      const headerCell = context.workbook.getActiveWorksheet().getCell(originalRange.rowIndex, colIndex);
      headerCell.values = [[suggestion.header]];
      headerCell.format.font.bold = true;
      headerCell.format.fill.color = "#E7E6E6";
      
      // Add formula or placeholder
      if (suggestion.formula && suggestion.formula.startsWith('=')) {
        const firstDataCell = context.workbook.getActiveWorksheet().getCell(originalRange.rowIndex + 1, colIndex);
        firstDataCell.formulas = [[suggestion.formula]];
        
        if (originalRange.rowCount > 2) {
          const copyRange = context.workbook.getActiveWorksheet().getRange(
            originalRange.rowIndex + 2, colIndex, originalRange.rowCount - 2, 1
          );
          firstDataCell.copyTo(copyRange);
        }
      }
    }
    
    await context.sync();
  } catch (error) {
    console.error("Error in applySimpleAnalysis:", error);
  }
}

async function applyFormulaToSpreadsheet(context, originalRange, response) {
  try {
    const formulaMatch = response.match(/FORMULA:\s*(.+)/);
    const headerMatch = response.match(/HEADER:\s*(.+)/);
    
    if (!formulaMatch) {
      showStatus("No formula found in response", "error");
      return;
    }
    
    const formula = formulaMatch[1].trim();
    const header = headerMatch ? headerMatch[1].trim() : "Result";
    
    const nextCol = originalRange.columnIndex + originalRange.columnCount;
    
    // Add header
    const headerCell = context.workbook.getActiveWorksheet().getCell(originalRange.rowIndex, nextCol);
    headerCell.values = [[header]];
    headerCell.format.font.bold = true;
    headerCell.format.fill.color = "#E7E6E6";
    
    // Apply formula
    const firstDataCell = context.workbook.getActiveWorksheet().getCell(originalRange.rowIndex + 1, nextCol);
    firstDataCell.formulas = [[formula]];
    
    if (originalRange.rowCount > 2) {
      const copyRange = context.workbook.getActiveWorksheet().getRange(
        originalRange.rowIndex + 2, nextCol, originalRange.rowCount - 2, 1
      );
      firstDataCell.copyTo(copyRange);
    }
    
    await context.sync();
  } catch (error) {
    console.error("Error applying formula:", error);
  }
}

// Stub functions
async function autoCategorizeData() {
  showStatus("Auto-categorize feature coming soon!", "info");
}

async function cleanAndStandardizeData() {
  showStatus("Data cleaning feature coming soon!", "info");
}

async function addCalculations() {
  showStatus("Add calculations feature coming soon!", "info");
}

// Utility functions
async function insertResults() {
  try {
    if (!currentResults) {
      showStatus("No results to insert", "error");
      return;
    }
    
    await Excel.run(async (context) => {
      const sheets = context.workbook.worksheets;
      const timestamp = new Date().toISOString().slice(0, 16).replace(/[:-]/g, "");
      const sheetName = "Claude_" + timestamp;
      
      const newSheet = sheets.add(sheetName);
      const range = newSheet.getRange("A1");
      
      range.values = [[currentResults]];
      range.format.wrapText = true;
      range.format.verticalAlignment = "Top";
      
      newSheet.getRange("A:A").format.columnWidth = 80;
      newSheet.activate();
      
      await context.sync();
      
      showStatus("Results inserted in new sheet!", "success");
    });
  } catch (error) {
    showStatus("Error inserting results: " + error.message, "error");
  }
}

async function callClaudeAPI(prompt) {
  const apiKey = document.getElementById("api-key").value.trim();
  
  if (!apiKey) {
    throw new Error("Please enter your Anthropic API key");
  }
  
  const response = await fetch("https://localhost:3001/api/claude", {
    method: "POST",
    headers: {
      "Content-Type": "application/json"
    },
    body: JSON.stringify({
      apiKey: apiKey,
      prompt: prompt
    })
  });
  
  if (!response.ok) {
    const errorData = await response.json().catch(() => ({ error: { message: response.statusText } }));
    throw new Error("API Error: " + (errorData.error && errorData.error.message ? errorData.error.message : response.statusText));
  }
  
  const data = await response.json();
  return data.content;
}

function convertRangeToText(values) {
  return values.map(row => row.map(cell => cell || "").join("\t")).join("\n");
}

function parseColumnSuggestions(response) {
  const suggestions = [];
  const lines = response.split('\n');
  
  let currentSuggestion = {};
  
  for (let line of lines) {
    line = line.trim();
    
    if (line.startsWith('COLUMN:')) {
      if (currentSuggestion.header) {
        suggestions.push(currentSuggestion);
      }
      currentSuggestion = {
        header: line.replace('COLUMN:', '').trim()
      };
    } else if (line.startsWith('FORMULA:')) {
      currentSuggestion.formula = line.replace('FORMULA:', '').trim();
    } else if (line.startsWith('EXPLANATION:')) {
      currentSuggestion.explanation = line.replace('EXPLANATION:', '').trim();
    }
  }
  
  if (currentSuggestion.header) {
    suggestions.push(currentSuggestion);
  }
  
  return suggestions;
}

function displayResults(results) {
  currentResults = results;
  document.getElementById("results").textContent = results;
  document.getElementById("insert-results").disabled = false;
}

function showStatus(message, type) {
  const statusDiv = document.getElementById("status");
  statusDiv.textContent = message;
  statusDiv.className = "status " + type;
  statusDiv.style.display = "block";
  
  if (type === "success" || type === "info") {
    setTimeout(() => {
      statusDiv.style.display = "none";
    }, 3000);
  }
}

function showLoading(show) {
  document.getElementById("loading").style.display = show ? "block" : "none";
  
  const buttons = document.querySelectorAll(".ms-Button");
  buttons.forEach(button => {
    button.disabled = show;
  });
}