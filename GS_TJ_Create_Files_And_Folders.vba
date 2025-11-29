// ========================================
// TJ PROJECT - 12 MONTHS VERSION
// CONFIGURATION
// ========================================
const MAJALIS_FILES_FOLDER_NAME = "Majalis_Files";
const REGIONS_FILES_FOLDER_NAME = "Regions_Files";
const PASSWORD = "";

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
// MAIN FUNCTION - CREATE MAJLIS FILES
// ========================================
function createRegionMajlisFiles() {
  try {
    Logger.log("🚀 Starting TJ process...");
    
    const sourceSpreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    const sourceSheet = sourceSpreadsheet.getActiveSheet();
    const parentFolder = DriveApp.getFileById(sourceSpreadsheet.getId()).getParents().next();
    
    Logger.log("✅ Source: " + sourceSpreadsheet.getName());
    Logger.log("✅ Parent Folder: " + parentFolder.getName());
    
    const regionsFolder = getOrCreateFolder(parentFolder, MAJALIS_FILES_FOLDER_NAME);
    Logger.log("✅ Majalis_Files folder ready");
    
    const data = sourceSheet.getDataRange().getValues();
    const headers = data[0];
    const rows = data.slice(1);
    
    const colRegion = 0;
    const colMajlis = 1;
    const colTanziem = 2;
    const colID = 3;
    const colName = 6;
    
    const uniqueRegions = [...new Set(rows.map(row => row[colRegion]).filter(r => r))];
    Logger.log("📊 Unique Regions: " + uniqueRegions.length);
    
    for (let region of uniqueRegions) {
      Logger.log("\n📁 Processing Region: " + region);
      
      const regionFolder = getOrCreateFolder(regionsFolder, region);
      
      const regionRows = rows.filter(row => row[colRegion] === region);
      const uniqueMajlis = [...new Set(regionRows.map(row => row[colMajlis]).filter(m => m))];
      
      Logger.log("  📄 Majlis count: " + uniqueMajlis.length);
      
      for (let majlis of uniqueMajlis) {
        Logger.log("    ➡️ Creating: " + majlis);
        
        const filteredRows = rows.filter(row => 
          row[colRegion] === region && row[colMajlis] === majlis
        );
        
        const newSpreadsheet = SpreadsheetApp.create(majlis);
        const newSheet = newSpreadsheet.getActiveSheet();
        
        const newFile = DriveApp.getFileById(newSpreadsheet.getId());
        regionFolder.addFile(newFile);
        DriveApp.getRootFolder().removeFile(newFile);
        
        setupMajlisSheet(newSheet, filteredRows, colRegion, colMajlis, colTanziem, colID, colName);
        
        Logger.log("    ✅ Created: " + majlis);
      }
    }
    
    Logger.log("\n🎉 TJ PROCESS COMPLETED SUCCESSFULLY!");
    SpreadsheetApp.getUi().alert("✅ Success!\n\nAll Region folders and Majlis files have been created in the '" + MAJALIS_FILES_FOLDER_NAME + "' folder.");
    
  } catch (error) {
    Logger.log("❌ ERROR: " + error.toString());
    SpreadsheetApp.getUi().alert("❌ Error: " + error.toString());
  }
}

// ========================================
// SETUP INDIVIDUAL MAJLIS SHEET
// ========================================
function setupMajlisSheet(sheet, filteredRows, colRegion, colMajlis, colTanziem, colID, colName) {
  
  sheet.setName("Data");
  
  const headers = [
    "Region", "Majlis", "Tanziem", "ID", "Name", 
    "Budget", "Nov", "Dec", "Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct",
    "Bezahlt", "Rest", "Prozent"
  ];
  
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  sheet.getRange(1, 1, 1, headers.length).setFontWeight("bold");
  
  if (filteredRows.length > 0) {
    const outputData = filteredRows.map(row => [
      row[colRegion],
      row[colMajlis],
      row[colTanziem],
      row[colID],
      row[colName],
      "",
      "", "", "", "", "", "", "", "", "", "", "", "",
      "",
      "",
      ""
    ]);
    
    sheet.getRange(2, 1, outputData.length, outputData[0].length).setValues(outputData);
    
    const numRows = filteredRows.length;
    
    for (let i = 2; i <= numRows + 1; i++) {
      sheet.getRange(i, 18).setFormula(`=SUM(G${i}:Q${i})`);
      sheet.getRange(i, 19).setFormula(`=F${i}-R${i}`);
      sheet.getRange(i, 20).setFormula(`=IF(F${i}=0,0,R${i}/F${i})`);
    }
    
    sheet.getRange(2, 6, numRows, 14).setNumberFormat("#,##0.00");
    sheet.getRange(2, 20, numRows, 1).setNumberFormat("0.00%");
  }
  
  protectColumns(sheet);
  sheet.autoResizeColumns(1, headers.length);
}

// ========================================
// PROTECT COLUMNS
// ========================================
function protectColumns(sheet) {
  try {
    const protection1 = sheet.getRange("A:E").protect();
    protection1.setDescription("🔒 Protected: Region, Majlis, Tanziem, ID, Name");
    protection1.setWarningOnly(false);
    
    const protection2 = sheet.getRange("R:T").protect();
    protection2.setDescription("🔒 Protected: Bezahlt, Rest, Prozent (Formulas)");
    protection2.setWarningOnly(false);
    
    Logger.log("✅ Protected columns: A-E and R-T");
    
  } catch (error) {
    Logger.log("⚠️ Protection warning: " + error.toString());
  }
}

// ========================================
// CREATE REGION FILES
// ========================================
function createRegionFiles() {
  try {
    Logger.log("🚀 Starting Region Files creation...");
    
    const sourceSpreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    const sourceSheet = sourceSpreadsheet.getActiveSheet();
    const parentFolder = DriveApp.getFileById(sourceSpreadsheet.getId()).getParents().next();
    
    Logger.log("✅ Source: " + sourceSpreadsheet.getName());
    
    const regionsFilesFolder = getOrCreateFolder(parentFolder, REGIONS_FILES_FOLDER_NAME);
    Logger.log("✅ Regions_Files folder ready");
    
    const data = sourceSheet.getDataRange().getValues();
    const rows = data.slice(1);
    
    const uniqueRegions = [...new Set(rows.map(row => row[0]).filter(r => r))];
    Logger.log("📊 Unique Regions: " + uniqueRegions.length);
    
    for (let region of uniqueRegions) {
      Logger.log("\n📁 Creating file for: " + region);
      
      const regionSpreadsheet = SpreadsheetApp.create(region);
      const regionSheet = regionSpreadsheet.getActiveSheet();
      regionSheet.setName("Data");
      
      const regionFile = DriveApp.getFileById(regionSpreadsheet.getId());
      regionsFilesFolder.addFile(regionFile);
      DriveApp.getRootFolder().removeFile(regionFile);
      
      setupRegionSheet(regionSheet);
      
      Logger.log("✅ Created: " + region);
    }
    
    Logger.log("\n🎉 ALL REGION FILES CREATED!");
    SpreadsheetApp.getUi().alert(
      "✅ Success!\n\n" +
      "Created " + uniqueRegions.length + " Region files in '" + REGIONS_FILES_FOLDER_NAME + "' folder.\n\n" +
      "Next step: Add the collection script to each Region file.\n" +
      "(You can copy from one file to all others)"
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
  
  sheet.getRange(2, 1).setValue("Add script via Extensions → Apps Script, then click menu to refresh data");
  sheet.getRange(2, 1).setFontStyle("italic").setFontColor("#999999");
}

// ========================================
// CREATE MENU
// ========================================
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('🔧 TJ Project')
    .addItem('📁 Create Majlis Files', 'createRegionMajlisFiles')
    .addItem('📂 Create Region Files', 'createRegionFiles')
    .addToUi();
}
