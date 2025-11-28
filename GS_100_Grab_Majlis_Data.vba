// ========================================
// 100 DASHBOARD - DATA COLLECTION
// CONFIGURATION
// ========================================
const MAJALIS_DATA_SHEET_NAME = "Majalis_Data";  // ⚠️ Change if needed

// ========================================
// MAIN COLLECTION FUNCTION
// ========================================
function collectDataFromAllFiles() {
  try {
    Logger.log("🚀 Starting data collection...");
    
    // STEP 1: Get current file and parent folder
    const dashboardSpreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    const parentFolder = DriveApp.getFileById(dashboardSpreadsheet.getId()).getParents().next();
    
    Logger.log("✅ Dashboard: " + dashboardSpreadsheet.getName());
    Logger.log("✅ Parent Folder: " + parentFolder.getName());
    
    // STEP 2: Find "Regions" folder
    const regionsFolders = parentFolder.getFoldersByName("Regions");
    if (!regionsFolders.hasNext()) {
      throw new Error("❌ 'Regions' folder not found! Please run the file creation script first.");
    }
    const regionsFolder = regionsFolders.next();
    Logger.log("✅ Found Regions folder");
    
    // STEP 3: Prepare Dashboard sheet
    let dataSheet = dashboardSpreadsheet.getSheetByName(MAJALIS_DATA_SHEET_NAME);
    if (!dataSheet) {
      dataSheet = dashboardSpreadsheet.insertSheet(MAJALIS_DATA_SHEET_NAME);
      Logger.log("✅ Created new sheet: " + MAJALIS_DATA_SHEET_NAME);
    } else {
      dataSheet.clear();
      Logger.log("✅ Cleared existing sheet: " + MAJALIS_DATA_SHEET_NAME);
    }
    
    // STEP 4: Write headers (8 months: Jul-Nov, Dec, Jan, Feb, Mar, Apr, May, Jun)
    const headers = [
      "Region", "Majlis", "Tanziem", "Anzahl", "Nicht-Zahler",
      "Budget", "Jul-Nov", "Dec", "Jan", "Feb", "Mar", "Apr", "May", "Jun",
      "Bezahlt", "Rest", "Prozent"
    ];
    dataSheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    dataSheet.getRange(1, 1, 1, headers.length).setFontWeight("bold").setBackground("#FF6D00").setFontColor("#FFFFFF");
    
    // STEP 5: Collect data from all files
    const collectedData = [];
    let fileCount = 0;
    
    // Loop through Region folders
    const regionFolders = regionsFolder.getFolders();
    while (regionFolders.hasNext()) {
      const regionFolder = regionFolders.next();
      const regionName = regionFolder.getName();
      Logger.log("\n📁 Processing Region: " + regionName);
      
      // Loop through Majlis files in this region
      const majlisFiles = regionFolder.getFilesByType(MimeType.GOOGLE_SHEETS);
      while (majlisFiles.hasNext()) {
        const majlisFile = majlisFiles.next();
        const majlisName = majlisFile.getName();
        fileCount++;
        
        Logger.log("  📄 Processing file " + fileCount + ": " + majlisName);
        
        try {
          // Open the Majlis spreadsheet
          const majlisSpreadsheet = SpreadsheetApp.openById(majlisFile.getId());
          const majlisSheet = majlisSpreadsheet.getSheetByName("Data");
          
          if (!majlisSheet) {
            Logger.log("    ⚠️ 'Data' sheet not found, skipping...");
            continue;
          }
          
          // Get all data from the sheet
          const data = majlisSheet.getDataRange().getValues();
          if (data.length <= 1) {
            Logger.log("    ⚠️ No data rows, skipping...");
            continue;
          }
          
          // Remove header row
          const rows = data.slice(1);
          
          // Column indexes (0-based) for 8 months version
          const colRegion = 0;    // A
          const colMajlis = 1;    // B
          const colTanziem = 2;   // C
          const colBudget = 5;    // F
          const colJulNov = 6;    // G (first month column)
          const colBezahlt = 14;  // O
          const colRest = 15;     // P
          
          // Get unique Tanziem values
          const uniqueTanziem = [...new Set(rows.map(row => row[colTanziem]).filter(t => t))];
          Logger.log("    📊 Unique Tanziem: " + uniqueTanziem.length);
          
          // Process each Tanziem group
          for (let tanziem of uniqueTanziem) {
            // Filter rows for this Tanziem
            const tanziemRows = rows.filter(row => row[colTanziem] === tanziem);
            
            // Calculate Anzahl (count)
            const anzahl = tanziemRows.length;
            
            // Calculate Nicht-Zahler (Budget empty OR <1 OR not a number)
            const nichtZahler = tanziemRows.filter(row => {
              const budget = row[colBudget];
              return budget === "" || budget === null || budget === undefined || 
                     typeof budget !== "number" || budget < 1;
            }).length;
            
            // Sum Budget
            const sumBudget = tanziemRows.reduce((sum, row) => {
              const val = row[colBudget];
              return sum + (typeof val === "number" ? val : 0);
            }, 0);
            
            // Sum month columns (Jul-Nov to Jun: columns 6-13, 8 months)
            const sumMonths = [];
            for (let monthCol = 6; monthCol <= 13; monthCol++) {
              const sumMonth = tanziemRows.reduce((sum, row) => {
                const val = row[monthCol];
                return sum + (typeof val === "number" ? val : 0);
              }, 0);
              sumMonths.push(sumMonth);
            }
            
            // Sum Bezahlt
            const sumBezahlt = tanziemRows.reduce((sum, row) => {
              const val = row[colBezahlt];
              return sum + (typeof val === "number" ? val : 0);
            }, 0);
            
            // Sum Rest
            const sumRest = tanziemRows.reduce((sum, row) => {
              const val = row[colRest];
              return sum + (typeof val === "number" ? val : 0);
            }, 0);
            
            // Calculate Prozent (Bezahlt / Budget)
            const prozent = sumBudget > 0 ? sumBezahlt / sumBudget : 0;
            
            // Add row to collected data
            collectedData.push([
              regionName,
              majlisName,
              tanziem,
              anzahl,
              nichtZahler,
              sumBudget,
              ...sumMonths,  // Jul-Nov, Dec, Jan, Feb, Mar, Apr, May, Jun (8 months)
              sumBezahlt,
              sumRest,
              prozent
            ]);
            
            Logger.log("      ✅ " + tanziem + ": " + anzahl + " rows, " + nichtZahler + " nicht-zahler");
          }
          
        } catch (fileError) {
          Logger.log("    ❌ Error processing file: " + fileError.toString());
        }
      }
    }
    
    // STEP 6: Write all collected data
    if (collectedData.length > 0) {
      dataSheet.getRange(2, 1, collectedData.length, collectedData[0].length).setValues(collectedData);
      
      // Format Budget and month columns as numbers (columns 6-16: Budget through Rest)
      dataSheet.getRange(2, 6, collectedData.length, 11).setNumberFormat("#,##0.00");
      
      // Format Prozent as percentage (column 17)
      dataSheet.getRange(2, 17, collectedData.length, 1).setNumberFormat("0.00%");
      
      // Auto-resize columns
      dataSheet.autoResizeColumns(1, headers.length);
      
      Logger.log("\n✅ Written " + collectedData.length + " rows to dashboard");
    } else {
      Logger.log("\n⚠️ No data collected");
    }
    
    Logger.log("\n🎉 DATA COLLECTION COMPLETED!");
    Logger.log("📊 Processed " + fileCount + " files");
    Logger.log("📝 Collected " + collectedData.length + " data rows");
    
    SpreadsheetApp.getUi().alert(
      "✅ Success!\n\n" +
      "Processed: " + fileCount + " files\n" +
      "Collected: " + collectedData.length + " rows\n\n" +
      "Data updated in '" + MAJALIS_DATA_SHEET_NAME + "' sheet."
    );
    
  } catch (error) {
    Logger.log("❌ ERROR: " + error.toString());
    SpreadsheetApp.getUi().alert("❌ Error:\n\n" + error.toString());
  }
}

// ========================================
// CREATE MENU
// ========================================
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('🔧 100 Dashboard')
    .addItem('🔄 Collect Data from All Files', 'collectDataFromAllFiles')
    .addToUi();
}
