// ========================================
// REGION DATA COLLECTION - TJ VERSION
// CONFIGURATION
// ========================================
const PASSWORD = "test";  // ⚠️ CHANGE THIS!
const MAJALIS_FILES_FOLDER_NAME = "Majalis_Files";

// ========================================
// REFRESH REGION DATA
// ========================================
function refreshRegionData() {
  try {
    const currentSpreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    const currentFileName = currentSpreadsheet.getName(); // Auto-detect region name
    
    Logger.log("🚀 Refreshing data for: " + currentFileName);
    
    const dataSheet = currentSpreadsheet.getSheetByName("Data");
    if (!dataSheet) {
      throw new Error("❌ 'Data' sheet not found!");
    }
    
    // Remove existing protection temporarily
    const protections = dataSheet.getProtections(SpreadsheetApp.ProtectionType.SHEET);
    protections.forEach(p => p.remove());
    
    // Clear existing data (keep headers)
    if (dataSheet.getLastRow() > 1) {
      dataSheet.getRange(2, 1, dataSheet.getLastRow() - 1, dataSheet.getLastColumn()).clear();
    }
    
    // Find parent folder structure
    const currentFile = DriveApp.getFileById(currentSpreadsheet.getId());
    const parentFolder = currentFile.getParents().next(); // Regions_Files folder
    const rootFolder = parentFolder.getParents().next();  // TJ_Project folder
    
    // Find Majalis_Files folder
    const majalisFilesFolders = rootFolder.getFoldersByName(MAJALIS_FILES_FOLDER_NAME);
    if (!majalisFilesFolders.hasNext()) {
      throw new Error("❌ '" + MAJALIS_FILES_FOLDER_NAME + "' folder not found!");
    }
    const majalisFilesFolder = majalisFilesFolders.next();
    
    // Find this region's folder (same name as file)
    const regionFolders = majalisFilesFolder.getFoldersByName(currentFileName);
    if (!regionFolders.hasNext()) {
      throw new Error("❌ '" + currentFileName + "' folder not found in Majalis_Files!");
    }
    const regionFolder = regionFolders.next();
    
    Logger.log("✅ Found folder: " + currentFileName);
    
    // Collect data from all Majlis files
    const collectedData = [];
    let fileCount = 0;
    
    const majlisFiles = regionFolder.getFilesByType(MimeType.GOOGLE_SHEETS);
    while (majlisFiles.hasNext()) {
      const majlisFile = majlisFiles.next();
      const majlisName = majlisFile.getName();
      fileCount++;
      
      Logger.log("  📄 Processing: " + majlisName);
      
      try {
        const majlisSpreadsheet = SpreadsheetApp.openById(majlisFile.getId());
        const majlisSheet = majlisSpreadsheet.getSheetByName("Data");
        
        if (!majlisSheet) {
          Logger.log("    ⚠️ 'Data' sheet not found, skipping...");
          continue;
        }
        
        const data = majlisSheet.getDataRange().getValues();
        if (data.length <= 1) {
          Logger.log("    ⚠️ No data rows, skipping...");
          continue;
        }
        
        const rows = data.slice(1);
        
        // Column indexes
        const colMajlis = 1;    // B
        const colTanziem = 2;   // C
        const colBudget = 5;    // F
        const colBezahlt = 17;  // R
        const colRest = 18;     // S
        
        // Get unique Tanziem
        const uniqueTanziem = [...new Set(rows.map(row => row[colTanziem]).filter(t => t))];
        
        for (let tanziem of uniqueTanziem) {
          const tanziemRows = rows.filter(row => row[colTanziem] === tanziem);
          
          // Calculate Anzahl
          const anzahl = tanziemRows.length;
          
          // Calculate Nicht-Zahler
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
          
          // Sum months (columns 6-17: Nov-Oct)
          const sumMonths = [];
          for (let monthCol = 6; monthCol <= 17; monthCol++) {
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
          
          // Calculate Prozent
          const prozent = sumBudget > 0 ? sumBezahlt / sumBudget : 0;
          
          collectedData.push([
            majlisName,
            tanziem,
            anzahl,
            nichtZahler,
            sumBudget,
            ...sumMonths,
            sumBezahlt,
            sumRest,
            prozent
          ]);
        }
        
      } catch (fileError) {
        Logger.log("    ❌ Error: " + fileError.toString());
      }
    }
    
    // Write data
    if (collectedData.length > 0) {
      dataSheet.getRange(2, 1, collectedData.length, collectedData[0].length).setValues(collectedData);
      
      // Format Budget and months as numbers
      dataSheet.getRange(2, 5, collectedData.length, 13).setNumberFormat("#,##0.00");
      
      // Format Prozent as percentage
      dataSheet.getRange(2, 20, collectedData.length, 1).setNumberFormat("0.00%");
      
      dataSheet.autoResizeColumns(1, 20);
      
      Logger.log("✅ Written " + collectedData.length + " rows");
    } else {
      Logger.log("⚠️ No data collected");
    }
    
    // Re-protect the sheet
    const protection = dataSheet.protect();
    protection.setDescription("Protected: " + currentFileName + " Data");
    if (PASSWORD) {
      protection.setPassword(PASSWORD);
    }
    
    Logger.log("✅ Sheet re-protected");
    Logger.log("🎉 REFRESH COMPLETED!");
    
    SpreadsheetApp.getUi().alert(
      "✅ Data Refreshed!\n\n" +
      "Region: " + currentFileName + "\n" +
      "Files processed: " + fileCount + "\n" +
      "Rows collected: " + collectedData.length
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
  ui.createMenu('🔧 Region Data')
    .addItem('🔄 Refresh Data', 'refreshRegionData')
    .addToUi();
}
