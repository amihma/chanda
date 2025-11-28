// ========================================
// 100 DASHBOARD - DATA COLLECTION
// CONFIGURATION
// ========================================
const MAJALIS_DATA_SHEET_NAME = "Majalis_Data";
const REGION_TANZIEM_SHEET_NAME = "Region_Tanziem";

// ========================================
// VIEW 1: MAJLIS VIEW (Detailed)
// ========================================
function majlisView() {
  try {
    Logger.log("🚀 Starting Majlis View collection...");
    
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
    
    // STEP 4: Write headers (8 months)
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
          
          // Column indexes (0-based)
          const colRegion = 0;
          const colMajlis = 1;
          const colTanziem = 2;
          const colBudget = 5;
          const colJulNov = 6;
          const colBezahlt = 14;
          const colRest = 15;
          
          // Get unique Tanziem values
          const uniqueTanziem = [...new Set(rows.map(row => row[colTanziem]).filter(t => t))];
          Logger.log("    📊 Unique Tanziem: " + uniqueTanziem.length);
          
          // Process each Tanziem group
          for (let tanziem of uniqueTanziem) {
            // Filter rows for this Tanziem
            const tanziemRows = rows.filter(row => row[colTanziem] === tanziem);
            
            // Calculate Anzahl (count)
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
            
            // Sum month columns (8 months: columns 6-13)
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
            
            // Calculate Prozent
            const prozent = sumBudget > 0 ? sumBezahlt / sumBudget : 0;
            
            // Add row to collected data
            collectedData.push([
              regionName,
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
            
            Logger.log("      ✅ " + tanziem + ": " + anzahl + " rows");
          }
          
        } catch (fileError) {
          Logger.log("    ❌ Error processing file: " + fileError.toString());
        }
      }
    }
    
    // STEP 6: Write all collected data
    if (collectedData.length > 0) {
      dataSheet.getRange(2, 1, collectedData.length, collectedData[0].length).setValues(collectedData);
      
      // Format columns
      dataSheet.getRange(2, 6, collectedData.length, 11).setNumberFormat("#,##0.00");
      dataSheet.getRange(2, 17, collectedData.length, 1).setNumberFormat("0.00%");
      
      // Auto-resize columns
      dataSheet.autoResizeColumns(1, headers.length);
      
      Logger.log("\n✅ Written " + collectedData.length + " rows to Majalis_Data");
    }
    
    Logger.log("\n🎉 MAJLIS VIEW COMPLETED!");
    Logger.log("📊 Processed " + fileCount + " files");
    
    SpreadsheetApp.getUi().alert(
      "✅ Majlis View Complete!\n\n" +
      "Processed: " + fileCount + " files\n" +
      "Collected: " + collectedData.length + " rows"
    );
    
  } catch (error) {
    Logger.log("❌ ERROR: " + error.toString());
    SpreadsheetApp.getUi().alert("❌ Error:\n\n" + error.toString());
  }
}

// ========================================
// VIEW 2: REGION VIEW (Aggregated by Region & Tanziem)
// ========================================
function regionsView() {
  try {
    Logger.log("🚀 Starting Regions View...");
    
    const dashboardSpreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    const sourceSheet = dashboardSpreadsheet.getSheetByName(MAJALIS_DATA_SHEET_NAME);
    
    // Check if Majalis_Data exists
    if (!sourceSheet) {
      throw new Error("❌ 'Majalis_Data' sheet not found! Please run 'Majlis View' first.");
    }
    
    // Get data from Majalis_Data
    const sourceData = sourceSheet.getDataRange().getValues();
    if (sourceData.length <= 1) {
      throw new Error("❌ No data in 'Majalis_Data'. Please run 'Majlis View' first.");
    }
    
    const headers = sourceData[0];
    const rows = sourceData.slice(1);
    
    Logger.log("✅ Reading from Majalis_Data: " + rows.length + " rows");
    
    // Prepare Region_Tanziem sheet
    let regionSheet = dashboardSpreadsheet.getSheetByName(REGION_TANZIEM_SHEET_NAME);
    if (!regionSheet) {
      regionSheet = dashboardSpreadsheet.insertSheet(REGION_TANZIEM_SHEET_NAME);
      Logger.log("✅ Created new sheet: " + REGION_TANZIEM_SHEET_NAME);
    } else {
      regionSheet.clear();
      Logger.log("✅ Cleared existing sheet: " + REGION_TANZIEM_SHEET_NAME);
    }
    
    // Write headers (without "Majlis", 8 months)
    const newHeaders = [
      "Region", "Tanziem", "Anzahl", "Nicht-Zahler",
      "Budget", "Jul-Nov", "Dec", "Jan", "Feb", "Mar", "Apr", "May", "Jun",
      "Bezahlt", "Rest", "Prozent"
    ];
    regionSheet.getRange(1, 1, 1, newHeaders.length).setValues([newHeaders]);
    regionSheet.getRange(1, 1, 1, newHeaders.length).setFontWeight("bold").setBackground("#0F9D58").setFontColor("#FFFFFF");
    
    // Aggregate data by Region + Tanziem
    const aggregated = {};
    
    for (let row of rows) {
      const region = row[0];
      const majlis = row[1];      // Skip in output
      const tanziem = row[2];
      const anzahl = row[3];
      const nichtZahler = row[4];
      const budget = row[5];
      const months = row.slice(6, 14);  // 8 months
      const bezahlt = row[14];
      const rest = row[15];
      
      // Create unique key
      const key = region + "|" + tanziem;
      
      if (!aggregated[key]) {
        aggregated[key] = {
          region: region,
          tanziem: tanziem,
          anzahl: 0,
          nichtZahler: 0,
          budget: 0,
          months: [0, 0, 0, 0, 0, 0, 0, 0],
          bezahlt: 0,
          rest: 0
        };
      }
      
      // Aggregate values
      aggregated[key].anzahl += anzahl;
      aggregated[key].nichtZahler += nichtZahler;
      aggregated[key].budget += (typeof budget === "number" ? budget : 0);
      
      for (let i = 0; i < 8; i++) {
        aggregated[key].months[i] += (typeof months[i] === "number" ? months[i] : 0);
      }
      
      aggregated[key].bezahlt += (typeof bezahlt === "number" ? bezahlt : 0);
      aggregated[key].rest += (typeof rest === "number" ? rest : 0);
    }
    
    // Convert to array and calculate Prozent
    const outputData = [];
    for (let key in aggregated) {
      const item = aggregated[key];
      const prozent = item.budget > 0 ? item.bezahlt / item.budget : 0;
      
      outputData.push([
        item.region,
        item.tanziem,
        item.anzahl,
        item.nichtZahler,
        item.budget,
        ...item.months,
        item.bezahlt,
        item.rest,
        prozent
      ]);
    }
    
    // Sort by Region, then Tanziem
    outputData.sort((a, b) => {
      if (a[0] !== b[0]) return a[0].localeCompare(b[0]);
      return a[1].localeCompare(b[1]);
    });
    
    // Write data
    if (outputData.length > 0) {
      regionSheet.getRange(2, 1, outputData.length, outputData[0].length).setValues(outputData);
      
      // Format columns
      regionSheet.getRange(2, 5, outputData.length, 10).setNumberFormat("#,##0.00");
      regionSheet.getRange(2, 16, outputData.length, 1).setNumberFormat("0.00%");
      
      // Auto-resize columns
      regionSheet.autoResizeColumns(1, newHeaders.length);
      
      Logger.log("✅ Written " + outputData.length + " aggregated rows");
    }
    
    Logger.log("\n🎉 REGIONS VIEW COMPLETED!");
    
    SpreadsheetApp.getUi().alert(
      "✅ Regions View Complete!\n\n" +
      "Aggregated: " + outputData.length + " Region-Tanziem combinations"
    );
    
  } catch (error) {
    Logger.log("❌ ERROR: " + error.toString());
    SpreadsheetApp.getUi().alert("❌ Error:\n\n" + error.toString());
  }
}

// ========================================
// VIEW 3: CENTRAL VIEW (Aggregated by Tanziem only)
// ========================================
function centralView() {
  try {
    Logger.log("🚀 Starting Central View...");
    
    const dashboardSpreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    const sourceSheet = dashboardSpreadsheet.getSheetByName(MAJALIS_DATA_SHEET_NAME);
    
    // Check if Majalis_Data exists
    if (!sourceSheet) {
      throw new Error("❌ 'Majalis_Data' sheet not found! Please run 'Majlis View' first.");
    }
    
    // Get data from Majalis_Data
    const sourceData = sourceSheet.getDataRange().getValues();
    if (sourceData.length <= 1) {
      throw new Error("❌ No data in 'Majalis_Data'. Please run 'Majlis View' first.");
    }
    
    const headers = sourceData[0];
    const rows = sourceData.slice(1);
    
    Logger.log("✅ Reading from Majalis_Data: " + rows.length + " rows");
    
    // Prepare Central sheet
    const CENTRAL_SHEET_NAME = "Central_Data";
    let centralSheet = dashboardSpreadsheet.getSheetByName(CENTRAL_SHEET_NAME);
    if (!centralSheet) {
      centralSheet = dashboardSpreadsheet.insertSheet(CENTRAL_SHEET_NAME);
      Logger.log("✅ Created new sheet: " + CENTRAL_SHEET_NAME);
    } else {
      centralSheet.clear();
      Logger.log("✅ Cleared existing sheet: " + CENTRAL_SHEET_NAME);
    }
    
    // Write headers
    const newHeaders = [
      "Tanziem", "Anzahl", "mit Budget", "Nicht-Zahler", "Budget", "Bezahlt", "Rest"
    ];
    centralSheet.getRange(1, 1, 1, newHeaders.length).setValues([newHeaders]);
    centralSheet.getRange(1, 1, 1, newHeaders.length).setFontWeight("bold").setBackground("#9C27B0").setFontColor("#FFFFFF");
    
    // Aggregate data by Tanziem only
    const aggregated = {};
    
    for (let row of rows) {
      const tanziem = row[2];         // Column C
      const anzahl = row[3];          // Column D
      const nichtZahler = row[4];     // Column E
      const budget = row[5];          // Column F
      const bezahlt = row[14];        // Column O (100: column 14)
      const rest = row[15];           // Column P (100: column 15)
      
      if (!aggregated[tanziem]) {
        aggregated[tanziem] = {
          anzahl: 0,
          nichtZahler: 0,
          budget: 0,
          bezahlt: 0,
          rest: 0
        };
      }
      
      // Aggregate values
      aggregated[tanziem].anzahl += anzahl;
      aggregated[tanziem].nichtZahler += nichtZahler;
      aggregated[tanziem].budget += (typeof budget === "number" ? budget : 0);
      aggregated[tanziem].bezahlt += (typeof bezahlt === "number" ? bezahlt : 0);
      aggregated[tanziem].rest += (typeof rest === "number" ? rest : 0);
    }
    
    // Convert to array
    const outputData = [];
    for (let tanziem in aggregated) {
      const item = aggregated[tanziem];
      const mitBudget = item.anzahl - item.nichtZahler;
      
      outputData.push([
        tanziem,
        item.anzahl,
        mitBudget,
        item.nichtZahler,
        item.budget,
        item.bezahlt,
        item.rest
      ]);
    }
    
    // Sort by Tanziem
    outputData.sort((a, b) => a[0].localeCompare(b[0]));
    
    // Write data
    if (outputData.length > 0) {
      centralSheet.getRange(2, 1, outputData.length, outputData[0].length).setValues(outputData);
      
      // Format Budget, Bezahlt, Rest as numbers
      centralSheet.getRange(2, 5, outputData.length, 3).setNumberFormat("#,##0.00");
      
      // Auto-resize columns
      centralSheet.autoResizeColumns(1, newHeaders.length);
      
      Logger.log("✅ Written " + outputData.length + " Tanziem rows");
    }
    
    Logger.log("\n🎉 CENTRAL VIEW COMPLETED!");
    
    SpreadsheetApp.getUi().alert(
      "✅ Central View Complete!\n\n" +
      "Aggregated: " + outputData.length + " Tanziem groups"
    );
    
  } catch (error) {
    Logger.log("❌ ERROR: " + error.toString());
    SpreadsheetApp.getUi().alert("❌ Error:\n\n" + error.toString());
  }
}

// ========================================
// REFRESH ALL VIEWS
// ========================================
function refreshAllViews() {
  majlisView();           // Scan files (slow)
  Utilities.sleep(1000);  
  regionsView();          // Transform data (fast)
  Utilities.sleep(500);
  centralView();          // Transform data (fast)
  
  SpreadsheetApp.getUi().alert(
    "✅ All Views Refreshed!\n\n" +
    "✓ Majalis_Data\n" +
    "✓ Region_Tanziem\n" +
    "✓ Central_Data\n\n" +
    "All sheets have been updated."
  );
}

// ========================================
// CREATE MENU
// ========================================
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('🔧 TJ Dashboard')  // or '🔧 100 Dashboard'
    .addItem('📋 Majlis View (Detailed)', 'majlisView')
    .addItem('📊 Regions View (by Region+Tanziem)', 'regionsView')
    .addItem('🎯 Central View (by Tanziem)', 'centralView')
    .addSeparator()
    .addItem('⚡ Refresh All Views', 'refreshAllViews')
    .addToUi();
}
