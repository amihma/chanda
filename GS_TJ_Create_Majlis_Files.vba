// ========================================
// TJ PROJECT - 12 MONTHS VERSION
// PASSWORD CONFIGURATION
// ========================================
const MAJALIS_FILES_FOLDER_NAME = "Majalis_Files";
const PASSWORD = "Password";  // ⚠️ CHANGE THIS!

// ========================================
// MAIN FUNCTION
// ========================================
function createRegionMajlisFiles() {
  try {
    Logger.log("🚀 Starting TJ process...");
    
    // STEP 1: Get current spreadsheet and parent folder
    const sourceSpreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    const sourceSheet = sourceSpreadsheet.getActiveSheet();
    const parentFolder = DriveApp.getFileById(sourceSpreadsheet.getId()).getParents().next();
    
    Logger.log("✅ Source: " + sourceSpreadsheet.getName());
    Logger.log("✅ Parent Folder: " + parentFolder.getName());
    
    // STEP 2: Create "Regions" folder
    const regionsFolder = getOrCreateFolder(parentFolder, MAJALIS_FILES_FOLDER_NAME);
    Logger.log("✅ Regions folder ready");
    
    // STEP 3: Get all data
    const data = sourceSheet.getDataRange().getValues();
    const headers = data[0];
    const rows = data.slice(1); // Remove header row
    
    // Find column indexes
    const colRegion = 0;  // Column A
    const colMajlis = 1;  // Column B
    const colTanziem = 2; // Column C
    const colID = 3;      // Column D
    const colName = 6;    // Column G
    
    // STEP 4: Get unique regions
    const uniqueRegions = [...new Set(rows.map(row => row[colRegion]).filter(r => r))];
    Logger.log("📊 Unique Regions: " + uniqueRegions.length);
    
    // STEP 5: Loop through each region
    for (let region of uniqueRegions) {
      Logger.log("\n📁 Processing Region: " + region);
      
      // Create region folder
      const regionFolder = getOrCreateFolder(regionsFolder, region);
      
      // Get unique Majlis for this region
      const regionRows = rows.filter(row => row[colRegion] === region);
      const uniqueMajlis = [...new Set(regionRows.map(row => row[colMajlis]).filter(m => m))];
      
      Logger.log("  📄 Majlis count: " + uniqueMajlis.length);
      
      // STEP 6: Loop through each Majlis
      for (let majlis of uniqueMajlis) {
        Logger.log("    ➡️ Creating: " + majlis);
        
        // Filter data for this Region + Majlis
        const filteredRows = rows.filter(row => 
          row[colRegion] === region && row[colMajlis] === majlis
        );
        
        // Create new spreadsheet
        const newSpreadsheet = SpreadsheetApp.create(majlis);
        const newSheet = newSpreadsheet.getActiveSheet();
        
        // Move file to region folder
        const newFile = DriveApp.getFileById(newSpreadsheet.getId());
        regionFolder.addFile(newFile);
        DriveApp.getRootFolder().removeFile(newFile);
        
        // Setup the new sheet
        setupMajlisSheet(newSheet, filteredRows, colRegion, colMajlis, colTanziem, colID, colName);
        
        Logger.log("    ✅ Created: " + majlis);
      }
    }
    
    Logger.log("\n🎉 TJ PROCESS COMPLETED SUCCESSFULLY!");
    SpreadsheetApp.getUi().alert("✅ Success!\n\nAll Region folders and Majlis files have been created in the 'Regions' folder.");
    
  } catch (error) {
    Logger.log("❌ ERROR: " + error.toString());
    SpreadsheetApp.getUi().alert("❌ Error: " + error.toString());
  }
}

// ========================================
// SETUP INDIVIDUAL MAJLIS SHEET
// ========================================
function setupMajlisSheet(sheet, filteredRows, colRegion, colMajlis, colTanziem, colID, colName) {
  
  // Rename sheet to "Data"
  sheet.setName("Data");
  
  // STEP 1: Create headers (12 months: Nov-Oct)
  const headers = [
    "Region", "Majlis", "Tanziem", "ID", "Name", 
    "Budget", "Nov", "Dec", "Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct",
    "Bezahlt", "Rest", "Prozent"
  ];
  
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  sheet.getRange(1, 1, 1, headers.length).setFontWeight("bold");
  
  // STEP 2: Write data (only columns A, B, C, D, G -> A, B, C, D, E)
  if (filteredRows.length > 0) {
    const outputData = filteredRows.map(row => [
      row[colRegion],  // A: Region
      row[colMajlis],  // B: Majlis
      row[colTanziem], // C: Tanziem
      row[colID],      // D: ID
      row[colName],    // E: Name
      "",              // F: Budget (empty)
      "", "", "", "", "", "", "", "", "", "", "", "", // G-R: Nov-Oct (empty)
      "",              // S: Bezahlt (formula will be added)
      "",              // T: Rest (formula will be added)
      ""               // U: Prozent (formula will be added)
    ]);
    
    sheet.getRange(2, 1, outputData.length, outputData[0].length).setValues(outputData);
    
    // STEP 3: Add formulas
    const numRows = filteredRows.length;
    
    for (let i = 2; i <= numRows + 1; i++) {
      // Column R (18): Bezahlt = SUM(G:Q) [Nov:Oct]
      sheet.getRange(i, 18).setFormula(`=SUM(G${i}:Q${i})`);
      
      // Column S (19): Rest = Budget - Bezahlt
      sheet.getRange(i, 19).setFormula(`=F${i}-R${i}`);
      
      // Column T (20): Prozent = Bezahlt/Budget
      sheet.getRange(i, 20).setFormula(`=IF(F${i}=0,0,R${i}/F${i})`);
    }
    
    // STEP 4: Format columns
    // F-S (6-19): Number format
    sheet.getRange(2, 6, numRows, 14).setNumberFormat("#,##0.00");
    
    // T (20): Percentage format
    sheet.getRange(2, 20, numRows, 1).setNumberFormat("0.00%");
  }
  
  // STEP 5: Protect columns
  protectColumns(sheet, PASSWORD);
  
  // STEP 6: Auto-resize columns
  sheet.autoResizeColumns(1, headers.length);
}

// ========================================
// PROTECT COLUMNS WITH PASSWORD
// ========================================
function protectColumns(sheet, password) {
  // Protect columns A-E (1-5)
  const protection1 = sheet.getRange("A:E").protect();
  protection1.setDescription("Protected: Region, Majlis, Tanziem, ID, Name");
  protection1.setWarningOnly(false);
  if (password) {
    protection1.setPassword(password);
  }
  
  // Protect columns R-T (18-20)
  const protection2 = sheet.getRange("R:T").protect();
  protection2.setDescription("Protected: Bezahlt, Rest, Prozent");
  protection2.setWarningOnly(false);
  if (password) {
    protection2.setPassword(password);
  }
}

// ========================================
// GET OR CREATE FOLDER
// ========================================
function getOrCreateFolder(parentFolder, folderName) {
  const folders = parentFolder.getFoldersByName(folderName);
  
  if (folders.hasNext()) {
    return folders.next();
  } else {
    return parentFolder.createFolder(folderName);
  }
}

// ========================================
// TJ PROJECT - REGION FILES CREATOR
// CONFIGURATION
// ========================================
const PASSWORD = "test";  // ⚠️ CHANGE THIS!
const MAJALIS_FILES_FOLDER_NAME = "Majalis_Files";
const REGIONS_FILES_FOLDER_NAME = "Regions_Files";

// ========================================
// CREATE REGION FILES
// ========================================
function createRegionFiles() {
  try {
    Logger.log("🚀 Starting Region Files creation...");
    
    // Get current file and parent folder
    const sourceSpreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    const sourceSheet = sourceSpreadsheet.getActiveSheet();
    const parentFolder = DriveApp.getFileById(sourceSpreadsheet.getId()).getParents().next();
    
    Logger.log("✅ Source: " + sourceSpreadsheet.getName());
    
    // Create "Regions_Files" folder
    const regionsFilesFolder = getOrCreateFolder(parentFolder, REGIONS_FILES_FOLDER_NAME);
    Logger.log("✅ Regions_Files folder ready");
    
    // Get all data
    const data = sourceSheet.getDataRange().getValues();
    const rows = data.slice(1); // Remove header
    
    // Get unique regions
    const uniqueRegions = [...new Set(rows.map(row => row[0]).filter(r => r))];
    Logger.log("📊 Unique Regions: " + uniqueRegions.length);
    
    // Create a file for each region
    for (let region of uniqueRegions) {
      Logger.log("\n📁 Creating file for: " + region);
      
      // Create new spreadsheet
      const regionSpreadsheet = SpreadsheetApp.create(region);
      const regionSheet = regionSpreadsheet.getActiveSheet();
      regionSheet.setName("Data");
      
      // Move to Regions_Files folder
      const regionFile = DriveApp.getFileById(regionSpreadsheet.getId());
      regionsFilesFolder.addFile(regionFile);
      DriveApp.getRootFolder().removeFile(regionFile);
      
      // Setup sheet structure
      setupRegionSheet(regionSheet);
      
      // Add collection script to the Region file
      addCollectionScriptToRegionFile(regionSpreadsheet, region);
      
      Logger.log("✅ Created: " + region);
    }
    
    Logger.log("\n🎉 ALL REGION FILES CREATED!");
    SpreadsheetApp.getUi().alert(
      "✅ Success!\n\n" +
      "Created " + uniqueRegions.length + " Region files in '" + REGIONS_FILES_FOLDER_NAME + "' folder.\n\n" +
      "Each file has its own data collection script."
    );
    
  } catch (error) {
    Logger.log("❌ ERROR: " + error.toString());
    SpreadsheetApp.getUi().alert("❌ Error:\n\n" + error.toString());
  }
}

// ========================================
// SETUP REGION SHEET STRUCTURE
// ========================================
function setupRegionSheet(sheet) {
  // Headers (NO Region column - TJ 12 months)
  const headers = [
    "Majlis", "Tanziem", "Anzahl", "Nicht-Zahler",
    "Budget", "Nov", "Dec", "Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct",
    "Bezahlt", "Rest", "Prozent"
  ];
  
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  sheet.getRange(1, 1, 1, headers.length)
    .setFontWeight("bold")
    .setBackground("#E91E63")
    .setFontColor("#FFFFFF");
  
  sheet.autoResizeColumns(1, headers.length);
  
  // Add placeholder text
  sheet.getRange(2, 1).setValue("Click menu: 🔧 Region Data → 🔄 Refresh Data");
}

// ========================================
// ADD COLLECTION SCRIPT TO REGION FILE
// ========================================
function addCollectionScriptToRegionFile(spreadsheet, regionName) {
  const scriptId = spreadsheet.getId();
  
  // Get the script project
  try {
    const scriptContent = getRegionCollectionScript(regionName);
    
    // Note: We'll provide the script as a string that users need to manually add
    // or we embed it as a note (Google Apps Script API required for auto-injection)
    
    // For now, add instruction in a separate sheet
    const instructionSheet = spreadsheet.insertSheet("README");
    instructionSheet.getRange(1, 1).setValue("📋 SETUP INSTRUCTIONS");
    instructionSheet.getRange(2, 1).setValue(
      "1. Go to Extensions → Apps Script\n" +
      "2. Delete existing code\n" +
      "3. Copy the script from 'Script' sheet\n" +
      "4. Paste and Save\n" +
      "5. Refresh this file - menu will appear\n" +
      "6. Click: 🔧 Region Data → 🔄 Refresh Data"
    );
    
    const scriptSheet = spreadsheet.insertSheet("Script");
    scriptSheet.getRange(1, 1).setValue(scriptContent);
    scriptSheet.getRange(1, 1).setWrap(true);
    scriptSheet.setColumnWidth(1, 800);
    
  } catch (error) {
    Logger.log("⚠️ Could not add script automatically: " + error.toString());
  }
}

// ========================================
// GENERATE REGION COLLECTION SCRIPT
// ========================================
function getRegionCollectionScript(regionName) {
  return `// ========================================
// REGION DATA COLLECTION - ${regionName}
// AUTO-GENERATED SCRIPT
// ========================================
const PASSWORD = "test";  // ⚠️ CHANGE THIS!
const MAJALIS_FILES_FOLDER_NAME = "Majalis_Files";
const REGION_FOLDER_NAME = "${regionName}";  // Auto-set to file name

// ========================================
// REFRESH REGION DATA
// ========================================
function refreshRegionData() {
  try {
    Logger.log("🚀 Refreshing data for: " + REGION_FOLDER_NAME);
    
    const currentSpreadsheet = SpreadsheetApp.getActiveSpreadsheet();
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
    
    // Find this region's folder
    const regionFolders = majalisFilesFolder.getFoldersByName(REGION_FOLDER_NAME);
    if (!regionFolders.hasNext()) {
      throw new Error("❌ '" + REGION_FOLDER_NAME + "' folder not found in Majalis_Files!");
    }
    const regionFolder = regionFolders.next();
    
    Logger.log("✅ Found folder: " + REGION_FOLDER_NAME);
    
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
    protection.setDescription("Protected: " + REGION_FOLDER_NAME + " Data");
    if (PASSWORD) {
      protection.setPassword(PASSWORD);
    }
    
    Logger.log("✅ Sheet re-protected");
    Logger.log("🎉 REFRESH COMPLETED!");
    
    SpreadsheetApp.getUi().alert(
      "✅ Data Refreshed!\\n\\n" +
      "Region: " + REGION_FOLDER_NAME + "\\n" +
      "Files processed: " + fileCount + "\\n" +
      "Rows collected: " + collectedData.length
    );
    
  } catch (error) {
    Logger.log("❌ ERROR: " + error.toString());
    SpreadsheetApp.getUi().alert("❌ Error:\\n\\n" + error.toString());
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
}`;
}

// ========================================
// HELPER: GET OR CREATE FOLDER
// ========================================
function getOrCreateFolder(parentFolder, folderName) {
  const folders = parentFolder.getFoldersByName(folderName);
  if (folders.hasNext()) {
    return folders.next();
  } else {
    return parentFolder.createFolder(folderName);
  }
}

// ========================================
// UPDATE MENU (Add to existing menu)
// ========================================
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('🔧 TJ Project')
    .addItem('📁 Create Majlis Files', 'createRegionMajlisFiles')
    .addItem('📂 Create Region Files', 'createRegionFiles')
    .addToUi();
}
