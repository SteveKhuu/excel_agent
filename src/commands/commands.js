/* global console, Excel, Office */

Office.onReady(() => {
  console.log("Commands.js loaded");
});

async function analyzeSelection(event) {
  try {
    await Excel.run(async (context) => {
      const range = context.workbook.getSelectedRange();
      range.load("values, address, rowCount, columnCount");
      
      await context.sync();
      
      if (range.rowCount === 1 && range.columnCount === 1) {
        // Show message to user
        Office.context.ui.displayDialogAsync(
          'https://localhost:3000/dialog.html?message=Please select a data range with multiple cells',
          { height: 200, width: 400 }
        );
        return;
      }
      
      // For quick analysis, create a simple summary
      const dataText = convertRangeToText(range.values);
      
      // Create results sheet
      const sheets = context.workbook.worksheets;
      const timestamp = new Date().toISOString().slice(0, 16).replace(/[:-]/g, "");
      const sheetName = `Quick_Analysis_${timestamp}`;
      
      const newSheet = sheets.add(sheetName);
      
      // Add summary info
      const summaryRange = newSheet.getRange("A1:B10");
      summaryRange.values = [
        ["Quick Data Analysis", ""],
        ["Range Analyzed:", range.address],
        ["Dimensions:", `${range.rowCount} rows × ${range.columnCount} columns`],
        ["Timestamp:", new Date().toLocaleString()],
        ["", ""],
        ["Sample Data:", ""],
        ...dataText.split('\n').slice(0, 5).map(row => [row, ""])
      ];
      
      // Format the sheet
      newSheet.getRange("A1").format.font.bold = true;
      newSheet.getRange("A1").format.font.size = 14;
      newSheet.getRange("A:A").format.columnWidth = 20;
      newSheet.getRange("B:B").format.columnWidth = 40;
      
      newSheet.activate();
      
      await context.sync();
      
      // Show completion message
      Office.context.ui.displayDialogAsync(
        'https://localhost:3000/dialog.html?message=Quick analysis complete! Check the new worksheet.',
        { height: 200, width: 400 }
      );
    });
  } catch (error) {
    console.error("Analysis error:", error);
    Office.context.ui.displayDialogAsync(
      `https://localhost:3000/dialog.html?message=Error: ${error.message}`,
      { height: 200, width: 400 }
    );
  }
  
  event.completed();
}

function convertRangeToText(values) {
  return values
    .map(row => row.map(cell => cell || "").join("\t"))
    .join("\n");
}

// Register the function
Office.actions.associate("analyzeSelection", analyzeSelection);