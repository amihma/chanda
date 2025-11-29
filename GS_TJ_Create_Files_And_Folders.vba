// ========================================
// TJ PROJECT - WITH SKIP CONTROL
// CONFIGURATION
// ========================================
const MAJALIS_FILES_FOLDER_NAME = "Majalis_Files";
const REGIONS_FILES_FOLDER_NAME = "Regions_Files";
const PASSWORD = [REDACTED:PASSWORD]6";

// ========================================
// USER CONTROLS (Update these before each run)
// ========================================
const IF_FOLDERS_CREATED = "No";  // "Yes" or "No"
const FILES_TO_SKIP = 0;          // Number of files already created (0, 50, 100, etc.)

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
// MAIN - CREATE MAJLIS FILES
// ========================================
function createRegionMajlisFiles() {
  try {
    const startTime = new Date();
    const sourceSpreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    const sourceSheet = sourceSpreadsheet.getActiveSheet();
    const parentFolder = DriveApp.getFileById(sourceSpreadsheet.getId()).getParents().next();
    
    Logger.log("🚀 Starting process...");
    Logger.log("⚙️ IF_FOLDERS_CREATED: " + IF_FOLDERS_CREATED);
    Logger.log("⚙️ FILES_TO_SKIP: " + FILES_TO_SKIP);
    
    const regionsFolder = getOrCreateFolder(parentFolder, MAJALIS_FILES_FOLDER_NAME);
    
    // STEP 1: Read only needed columns
    sourceSpreadsheet.toast("Reading data...", "⏳ Step 1/3", -1);
    const lastRow = sourceSheet.getLastRow();
    const colA = sourceSheet.getRange(2, 1, lastRow - 1, 1).getValues(); // Region
    const colB = sourceSheet.getRange(2, 2, lastRow - 1, 1).getValues(); // Majlis
    const colC = sourceSheet.getRange(2, 3, lastRow - 1, 1).getValues(); // Tanziem
    const colD = sourceSheet.getRange(2, 4, lastRow - 1, 1).getValues(); // ID
    const colG = sourceSheet.getRange(2, 7, lastRow - 1, 1).getValues(); // Name
    
    Logger.log("✅ Read " + (lastRow - 1) + " rows");
    
    // STEP 2: Build file list and data map
    sourceSpreadsheet.toast("Organizing data...", "⏳ Step 2/3", -1);
    const fileList = [];
    const dataMap = {};
    const regionSet = new Set();
    const seenFiles = new Set();
    
    for (let i = 0; i < colA.length; i++) {
      const region = colA[i][0];
      const majlis = colB[i][0];
      const tanziem = colC[i][0];
      const id = colD[i][0];
      const name = colG[i][0];
      
      if (!region || !majlis) continue;
      
      regionSet.add(region);
      
      const fileKey = region + "|" + majlis;
      
      // Add to file list (only once per unique Region+Majlis)
      if (!seenFiles.has(fileKey)) {
        fileList.push({region: region, majlis: majlis});
        seenFiles.add(fileKey);
      }
      
      // Store data
      if (!dataMap[fileKey]) {
        dataMap[fileKey] = [];
      }
      dataMap[fileKey].push([region, majlis, tanziem, id, name]);
    }
    
    const totalFiles = fileList.length;
    Logger.log("📊 Total files to create: " + totalFiles);
    Logger.log("📊 Unique regions: " + regionSet.size);
    
    // STEP 3: Create/Get region folders
    let regionFolders = {};
    
    if (IF_FOLDERS_CREATED === "No") {
      sourceSpreadsheet.toast("Creating region folders...", "⏳ Step 3/3", -1);
      Logger.log("📁 Creating region folders...");
      
      for (let region of regionSet) {
        const folder = getOrCreateFolder(regionsFolder, region);
        regionFolders[region] = folder;
        Logger.log("  ✅ Created/Found folder: " + region);
      }
    } else {
      sourceSpreadsheet.toast("Getting existing folders...", "⏳ Step 3/3", -1);
      Logger.log("📁 Getting existing region folders...");
      
      for (let region of regionSet) {
        const folders = regionsFolder.getFoldersByName(region);
        if (folders.hasNext()) {
          regionFolders[region] = folders.next();
          Logger.log("  ✅ Found folder: " + region);
        } else {
          throw new Error("Folder not found: " + region + ". Set IF_FOLDERS_CREATED='No'");
        }
      }
    }
    
    // STEP 4: Create files (with skip logic)
    if (FILES_TO_SKIP > 0) {
      Logger.log("⏩ Skipping first " + FILES_TO_SKIP + " files");
      sourceSpreadsheet.toast("Skipping first " + FILES_TO_SKIP + " files...", "⏳ Starting", 2);
    }
    
    let filesCreated = 0;
    
    for (let i = FILES_TO_SKIP; i < totalFiles; i++) {
      const file = fileList[i];
      const fileNumber = i + 1; // User-friendly numbering (1-based)
      
      if (fileNumber % 5 === 0 || fileNumber === totalFiles) {
        sourceSpreadsheet.toast(
          "Creating: " + file.majlis,
          "⏳ File " + fileNumber + "/" + totalFiles + " (" + Math.round((fileNumber/totalFiles)*100) + "%)",
          2
        );
      }
      
      Logger.log("📄 Creating file " + fileNumber + "/" + totalFiles + ": " + file.majlis + " (Region: " + file.region + ")");
      
      const fileKey = file.region + "|" + file.majlis;
      const data = dataMap[fileKey] || [];
      
      // Create spreadsheet
      const newSpreadsheet = SpreadsheetApp.create(file.majlis);
      const newSheet = newSpreadsheet.getActiveSheet();
      
      // Move to region folder
      const regionFolder = regionFolders[file.region];
      const newFile = DriveApp.getFileById(newSpreadsheet.getId());
      regionFolder.addFile(newFile);
      DriveApp.getRootFolder().removeFile(newFile);
      
      // Setup sheet
      setupMajlisSheet(newSheet, data);
      
      filesCreated++;
    }
    
    const endTime = new Date();
    const duration = ((endTime - startTime) / 1000).toFixed(1);
    
    Logger.log("\n✅ BATCH COMPLETED");
    Logger.log("📊 Files created this run: " + filesCreated);
    Logger.log("📊 Total progress: " + (FILES_TO_SKIP + filesCreated) + "/" + totalFiles);
    Logger.log("⏱️ Time: " + duration + " seconds");
    
    sourceSpreadsheet.toast("", "", 1);
    
    if (FILES_TO_SKIP + filesCreated >= totalFiles) {
      // ALL DONE
      SpreadsheetApp.getUi().alert(
        "✅ ALL FILES CREATED!\n\n" +
        "Total files: " + totalFiles + "\n" +
        "Time: " + duration + " seconds"
      );
    } else {
      // MORE TO DO
      const remaining = totalFiles - (FILES_TO_SKIP + filesCreated);
      SpreadsheetApp.getUi().alert(
        "⏸️ Batch Complete!\n\n" +
        "Created this run: " + filesCreated + " files\n" +
        "Total progress: " + (FILES_TO_SKIP + filesCreated) + "/" + totalFiles + "\n" +
        "Remaining: " + remaining + "\n" +
        "Time: " + duration + " seconds\n\n" +
        "To continue:\n" +
        "1. Update FILES_TO_SKIP = " + (FILES_TO_SKIP + filesCreated) + "\n" +
        "2. Update IF_FOLDERS_CREATED = \"Yes\"\n" +
        "3. Run again"
      );
    }
    
  } catch (error) {
    Logger.log("❌ ERROR: " + error.toString());
    SpreadsheetApp.getUi().alert("❌ Error:\n\n" + error.toString());
  }
}

// ========================================
// SETUP MAJLIS SHEET (CORRECTED COLUMNS)
// ========================================
function setupMajlisSheet(sheet, data) {
  sheet.setName("Data");
  
  // Headers: 21 columns total (A-U)
  const headers = [
    "Region", "Majlis", "Tanziem", "ID", "Name", 
    "Budget", "Nov", "Dec", "Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct",
    "Bezahlt", "Rest", "Prozent"
  ];
  
  sheet.getRange(1, 1, 1, 21).setValues([headers]);
  sheet.getRange(1, 1, 1, 21)
    .setFontWeight("bold")
    .setBackground("#4285F4")
    .setFontColor("#FFFFFF");
  
  if (data.length > 0) {
    // Data: [region, majlis, tanziem, id, name]
    const outputData = data.map(row => [
      row[0],  // A: Region
      row[1],  // B: Majlis
      row[2],  // C: Tanziem
      row[3],  // D: ID
      row[4],  // E: Name
      "",      // F: Budget (6)
      "", "", "", "", "", "", "", "", "", "", "", "",  // G-R: 12 months (7-18)
      "",      // S: Bezahlt (19)
      "",      // T: Rest (20)
      ""       // U: Prozent (21)
    ]);
    
    sheet.getRange(2, 1, outputData.length, 21).setValues(outputData);
    
    const numRows = outputData.length;
    
    // FORMULAS at columns S, T, U (19, 20, 21)
    const formulas = [];
    for (let i = 2; i <= numRows + 1; i++) {
      formulas.push([
        `=SUM(G${i}:R${i})`,           // S (19): Bezahlt = SUM of 12 months (G-R)
        `=F${i}-S${i}`,                // T (20): Rest = Budget - Bezahlt
        `=IF(F${i}=0,0,S${i}/F${i})`   // U (21): Prozent = Bezahlt/Budget
      ]);
    }
    
    sheet.getRange(2, 19, numRows, 3).setFormulas(formulas);
    
    // FORMATTING
    // Columns F-S (6-19): Number format for Budget and months
    sheet.getRange(2, 6, numRows, 14).setNumberFormat("#,##0.00");
    
    // Column U (21): Percentage format
    sheet.getRange(2, 21, numRows, 1).setNumberFormat("0.00%");
  }
  
  // PROTECTION
  try {
    const protection1 = sheet.getRange("A:E").protect();
    protection1.setDescription("🔒 Protected: Region, Majlis, Tanziem, ID, Name");
    protection1.setWarningOnly(false);
    
    const protection2 = sheet.getRange("S:U").protect();
    protection2.setDescription("🔒 Protected: Bezahlt, Rest, Prozent (Formulas)");
    protection2.setWarningOnly(false);
  } catch (e) {
    Logger.log("⚠️ Protection warning: " + e.toString());
  }
  
  // AUTO-RESIZE & FREEZE
  sheet.autoResizeColumns(1, 21);
  sheet.setFrozenRows(1);
}

// ========================================
// CREATE REGION FILES (Simple - No skip needed)
// ========================================
function createRegionFiles() {
  try {
    Logger.log("🚀 Starting Region Files creation...");
    
    const sourceSpreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    const sourceSheet = sourceSpreadsheet.getActiveSheet();
    const parentFolder = DriveApp.getFileById(sourceSpreadsheet.getId()).getParents().next();
    
    const regionsFilesFolder = getOrCreateFolder(parentFolder, REGIONS_FILES_FOLDER_NAME);
    
    // Read only column A (Region)
    const lastRow = sourceSheet.getLastRow();
    const colA = sourceSheet.getRange(2, 1, lastRow - 1, 1).getValues();
    const uniqueRegions = [...new Set(colA.map(row => row[0]).filter(r => r))];
    
    Logger.log("📊 Creating " + uniqueRegions.length + " Region files");
    
    for (let i = 0; i < uniqueRegions.length; i++) {
      const region = uniqueRegions[i];
      const fileNum = i + 1;
      
      sourceSpreadsheet.toast(
        "Creating: " + region,
        "⏳ " + fileNum + "/" + uniqueRegions.length,
        2
      );
      
      Logger.log("📄 Creating file " + fileNum + "/" + uniqueRegions.length + ": " + region);
      
      const regionSpreadsheet = SpreadsheetApp.create(region);
      const regionSheet = regionSpreadsheet.getActiveSheet();
      regionSheet.setName("Data");
      
      // Headers (20 columns - no Region column)
      const headers = [
        "Majlis", "Tanziem", "Anzahl", "Nicht-Zahler",
        "Budget", "Nov", "Dec", "Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct",
        "Bezahlt", "Rest", "Prozent"
      ];
      
      regionSheet.getRange(1, 1, 1, 20).setValues([headers]);
      regionSheet.getRange(1, 1, 1, 20)
        .setFontWeight("bold")
        .setBackground("#E91E63")
        .setFontColor("#FFFFFF");
      
      regionSheet.autoResizeColumns(1, 20);
      regionSheet.setFrozenRows(1);
      
      regionSheet.getRange(2, 1).setValue("Add script via Extensions → Apps Script to refresh data");
      regionSheet.getRange(2, 1).setFontStyle("italic").setFontColor("#999999");
      
      // Move to folder
      const regionFile = DriveApp.getFileById(regionSpreadsheet.getId());
      regionsFilesFolder.addFile(regionFile);
      DriveApp.getRootFolder().removeFile(regionFile);
    }
    
    Logger.log("✅ All Region files created");
    SpreadsheetApp.getUi().alert(
      "✅ Success!\n\n" +
      "Created " + uniqueRegions.length + " Region files in '" + REGIONS_FILES_FOLDER_NAME + "' folder."
    );
    
  } catch (error) {
    Logger.log("❌ ERROR: " + error.toString());
    SpreadsheetApp.getUi().alert("❌ Error:\n\n" + error.toString());
  }
}

// ========================================
// MENU
// ========================================
function onOpen() {
  SpreadsheetApp.getUi().createMenu('🔧 TJ Project')
    .addItem('📁 Create Majlis Files', 'createRegionMajlisFiles')
    .addItem('📂 Create Region Files', 'createRegionFiles')
    .addToUi();
}
