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
      const prompt = "Analyze this table data and create additional analysis tables that would provide business insights:\n\nSelected Data:\n" + dataText + "\n\nCreate 2-3 analysis tables in code blocks. For example:\n- Summary statistics table\n- Growth rate analysis\n- Performance categorization\n- Trend analysis\n- Comparative metrics\n\nFormat each analysis as a table inside ``` code blocks. Focus on actionable business insights derived from the original data.";

      const response = await callClaudeAPI(prompt);
      
      // Create analysis worksheet with the structured tables
      await createAnalysisWorksheet(context, response, range);
      
      displayResults(response);
      showStatus("Analysis tables created successfully!", "success");
    });
  } catch (error) {
    showStatus("Error: " + error.message, "error");
  } finally {
    showLoading(false);
  }
}

async function createFormula() {
  try {
    showLoading(true);
    
    const task = prompt("What type of analysis do you need? (e.g., 'growth rates', 'profit margins', 'variance analysis')");
    if (!task) return;
    
    await Excel.run(async (context) => {
      const range = context.workbook.getSelectedRange();
      range.load("values, address, rowCount, columnCount");
      
      await context.sync();
      
      const dataText = convertRangeToText(range.values);
      const prompt = "Create analysis tables for: " + task + "\n\nBased on this data:\n" + dataText + "\n\nGenerate tables with calculated results in code blocks. Include:\n- Formula explanations\n- Calculated results\n- Interpretation tables\n\nFormat as structured tables inside ``` code blocks.";

      const response = await callClaudeAPI(prompt);
      
      await createAnalysisWorksheet(context, response, range);
      
      displayResults(response);
      showStatus("Analysis tables created successfully!", "success");
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
      
      if (range.rowCount === 1 && range.columnCount === 1) {
        showStatus("Please select a larger data range", "error");
        return;
      }
      
      const dataText = convertRangeToText(range.values);
      const prompt = "Create business insight tables based on this data:\n\n" + dataText + "\n\nGenerate 2-3 insight tables in code blocks such as:\n- Key Performance Indicators\n- Risk Assessment Matrix\n- Opportunity Analysis\n- Benchmarking Comparisons\n- Action Items Priority Matrix\n\nFormat each insight as a structured table inside ``` code blocks.";

      const response = await callClaudeAPI(prompt);
      
      await createAnalysisWorksheet(context, response, range);
      
      displayResults(response);
      showStatus("Insight tables created successfully!", "success");
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
    const timestamp = new Date().toISOString().slice(0, 16).replace(/[:-]/g, "");
    const worksheetName = "Claude_" + timestamp;
    
    const newWorksheet = context.workbook.worksheets.add(worksheetName);
    
    const tables = extractOnlyCodeBlockTables(response);
    
    if (tables.length === 0) {
      showStatus("No code block tables found in response", "info");
      return;
    }
    
    let currentRow = 0;
    
    for (let table of tables) {
      currentRow = await insertCleanTable(context, newWorksheet, table, currentRow);
      currentRow += 2;
    }
    
    newWorksheet.getRange("A:A").format.columnWidth = 200;
    newWorksheet.getRange("B:F").format.columnWidth = 100;
    
    newWorksheet.activate();
    await context.sync();
    
  } catch (error) {
    showStatus("Error creating worksheet: " + error.message, "error");
  }
}

async function createAnalysisWorksheet(context, response, originalRange) {
  try {
    const timestamp = new Date().toISOString().slice(0, 16).replace(/[:-]/g, "");
    const worksheetName = "Analysis_" + timestamp;
    
    const newWorksheet = context.workbook.worksheets.add(worksheetName);
    
    // Add reference to original data
    let currentRow = 0;
    const headerCell = newWorksheet.getCell(currentRow, 0);
    headerCell.values = [["Analysis of: " + originalRange.address]];
    headerCell.format.font.bold = true;
    headerCell.format.font.size = 14;
    headerCell.format.fill.color = "#2F5496";
    headerCell.format.font.color = "white";
    currentRow += 2;
    
    // Extract and insert analysis tables
    const tables = extractOnlyCodeBlockTables(response);
    
    if (tables.length === 0) {
      // Fallback: show that no structured tables were found
      const noTablesCell = newWorksheet.getCell(currentRow, 0);
      noTablesCell.values = [["No structured analysis tables found in response."]];
      currentRow += 2;
      
      // Insert raw response as backup
      const lines = response.split('\n');
      for (let line of lines) {
        if (line.trim()) {
          const cell = newWorksheet.getCell(currentRow, 0);
          cell.values = [[line.trim()]];
          currentRow++;
        }
      }
    } else {
      // Insert structured analysis tables
      for (let table of tables) {
        currentRow = await insertCleanTable(context, newWorksheet, table, currentRow);
        currentRow += 2; // Add spacing between analysis tables
      }
    }
    
    // Format the worksheet
    newWorksheet.getRange("A:A").format.columnWidth = 200;
    newWorksheet.getRange("B:F").format.columnWidth = 100;
    
    newWorksheet.activate();
    await context.sync();
    
  } catch (error) {
    throw error;
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
    
    if (line.startsWith('```') || line === '```') {
      if (!inCodeBlock) {
        inCodeBlock = true;
        currentTable = [];
        tableTitle = findTableTitle(lines, i);
      } else {
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
      const cells = parseCleanRow(line);
      if (cells.length > 0) {
        currentTable.push(cells);
      }
    }
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
  
  let cleanValue = String(value).trim();
  
  // Handle Excel formula issues - replace problematic starting characters
  if (cleanValue === '+' || cleanValue === '-') {
    return '—'; // Use em dash for standalone + or -
  }
  
  // Handle cells that start with + or - but aren't numbers
  if (cleanValue.startsWith('+') && !cleanValue.match(/^\+[\d,\.]+/)) {
    cleanValue = '＋' + cleanValue.slice(1); // Use full-width plus
  }
  
  if (cleanValue.startsWith('-') && !cleanValue.match(/^-[\d,\.]+/)) {
    cleanValue = '—' + cleanValue.slice(1); // Use em dash
  }
  
  return cleanValue;
}

async function insertCleanTable(context, worksheet, table, startRow) {
  let currentRow = startRow;
  
  try {
    if (table.title && table.title !== "Data" && table.title !== "Table") {
      const titleCell = worksheet.getCell(currentRow, 0);
      titleCell.values = [[table.title]];
      titleCell.format.font.bold = true;
      titleCell.format.font.size = 12;
      titleCell.format.fill.color = "#4472C4";
      titleCell.format.font.color = "white";
      currentRow++;
    }
    
    for (let rowIndex = 0; rowIndex < table.rows.length; rowIndex++) {
      const row = table.rows[rowIndex];
      
      for (let colIndex = 0; colIndex < row.length; colIndex++) {
        const cell = worksheet.getCell(currentRow, colIndex);
        let cellValue = row[colIndex];
        
        if (cellValue === undefined || cellValue === null) {
          cell.values = [['']];
          continue;
        }
        
        // Clean the cell value to handle + and - issues
        cellValue = cleanCellValue(cellValue);
        
        // Handle numeric values
        const cleanNumValue = cellValue.replace(/,/g, '').replace(/[()$]/g, '');
        const numValue = parseFloat(cleanNumValue);
        
        if (!isNaN(numValue) && cellValue.match(/[\d,]/)) {
          cell.values = [[numValue]];
          
          if (Math.abs(numValue) >= 100) {
            cell.numberFormat = [["#,##0"]];
          }
          
          if (cellValue.includes('%')) {
            cell.numberFormat = [["0%"]];
            cell.values = [[numValue / 100]];
          }
          
          cell.format.horizontalAlignment = "Right";
        } else {
          cell.values = [[cellValue]];
          
          if (colIndex === 0) {
            cell.format.horizontalAlignment = "Left";
          }
        }
        
        if (rowIndex === 0 || cellValue.includes('Year') || cellValue.includes('$')) {
          cell.format.font.bold = true;
          cell.format.fill.color = "#D9E2F3";
        }
      }
      
      currentRow++;
    }
    
    return currentRow;
    
  } catch (error) {
    throw error;
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