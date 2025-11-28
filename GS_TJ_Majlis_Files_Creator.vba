// ========================================
// PASSWORD CONFIGURATION
// ========================================
const PASSWORD = "YourPassword123";  // ⚠️ CHANGE THIS!

// ========================================
// MAIN FUNCTION
// ========================================
function createRegionMajlisFiles() {
  try {
    Logger.log("🚀 Starting process...");
    
    // STEP 1: Get current spreadsheet and parent folder
    const sourceSpreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    const sourceSheet = sourceSpreadsheet.getActiveSheet();
    const parentFolder = DriveApp.getFileById(sourceSpreadsheet.getId()).getParents().next();
    
    Logger.log("✅ Source: " + sourceSpreadsheet.getName());
    Logger.log("✅ Parent Folder: " + parentFolder.getName());
    
    // STEP 2: Create "Regions" folder
    const regionsFolder = getOrCreateFolder(parentFolder, "Regions");
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
    
    Logger.log("\n🎉 PROCESS COMPLETED SUCCESSFULLY!");
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
  
  // STEP 1: Create headers
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
// CREATE MENU (Optional - for easy access)
// ========================================
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('🔧 TJ Project')
    .addItem('📁 Create Region/Majlis Files', 'createRegionMajlisFiles')
    .addToUi();
}
