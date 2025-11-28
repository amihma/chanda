// ========================================
// TJ PROJECT - EMAIL SYSTEM
// CONFIGURATION
// ========================================
// Sheet Names
const MAJALIS_DATA_SHEET = "Majalis_Data";
const REGION_DATA_SHEET = "Region_Tanziem";
const MAJLIS_EMAIL_SHEET = "Majlis_Email";
const REGION_EMAIL_SHEET = "Region_Email";

// Logo URL
const LOGO_URL = "https://your-domain.com/logo.png";  // ⚠️ UPDATE THIS

// Email Signature
const EMAIL_SIGNATURE = `Wasslam<br>
Muhammad Ahmad<br>
Qaid Umumi<br>
kontakt@example.com`;

// Sender Name
const SENDER_NAME = "TJ Admin";

// TJ Fiscal Year: November to October
const FISCAL_START_MONTH = 11; // November

// ========================================
// CALCULATE FISCAL YEAR (Nov-Oct)
// ========================================
function getFiscalYear() {
  const now = new Date();
  const currentYear = now.getFullYear();
  const currentMonth = now.getMonth() + 1; // 1-12
  
  if (currentMonth >= FISCAL_START_MONTH) {
    // Nov or Dec → 2025/26
    return currentYear + "/" + (currentYear + 1).toString().substr(2);
  } else {
    // Jan-Oct → 2024/25
    return (currentYear - 1) + "/" + currentYear.toString().substr(2);
  }
}

// ========================================
// ON EDIT TRIGGER (For Checkbox)
// ========================================
function onEdit(e) {
  if (!e) return;
  
  const sheet = e.source.getActiveSheet();
  const sheetName = sheet.getName();
  const range = e.range;
  const row = range.getRow();
  const col = range.getColumn();
  
  // Check if editing checkbox in column 3 (Send column)
  if (col !== 3 || row === 1) return; // Skip header
  
  const value = range.getValue();
  
  if (sheetName === MAJLIS_EMAIL_SHEET && value === true) {
    sendMajlisEmail(row);
    range.setValue(false); // Uncheck after sending
  } else if (sheetName === REGION_EMAIL_SHEET && value === true) {
    sendRegionEmail(row);
    range.setValue(false); // Uncheck after sending
  }
}

// ========================================
// SEND MAJLIS EMAIL
// ========================================
function sendMajlisEmail(row) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const emailSheet = ss.getSheetByName(MAJLIS_EMAIL_SHEET);
    const dataSheet = ss.getSheetByName(MAJALIS_DATA_SHEET);
    
    if (!emailSheet || !dataSheet) {
      throw new Error("Required sheets not found!");
    }
    
    // Get Majlis name and email
    const majlisName = emailSheet.getRange(row, 1).getValue();
    const emailAddress = emailSheet.getRange(row, 2).getValue();
    
    if (!emailAddress || emailAddress === "") {
      throw new Error("No email address found for " + majlisName);
    }
    
    Logger.log("Sending email to: " + majlisName + " (" + emailAddress + ")");
    
    // Get data from Majalis_Data
    const data = dataSheet.getDataRange().getValues();
    const rows = data.slice(1);
    
    // Find rows for this Majlis
    const majlisRows = rows.filter(row => row[1] === majlisName); // Col B = Majlis
    
    if (majlisRows.length === 0) {
      throw new Error("No data found for " + majlisName);
    }
    
    // Separate Khuddam and Atfal
    let khuddamData = null;
    let atfalData = null;
    
    for (let row of majlisRows) {
      const tanziem = row[2]; // Col C = Tanziem
      if (tanziem.toLowerCase().includes("khuddam")) {
        khuddamData = row;
      } else if (tanziem.toLowerCase().includes("atfal")) {
        atfalData = row;
      }
    }
    
    // Build and send email
    const fiscalYear = getFiscalYear();
    const htmlBody = buildMajlisEmailHTML(majlisName, khuddamData, atfalData, fiscalYear);
    
    MailApp.sendEmail({
      to: emailAddress,
      subject: "Chanda Summary " + fiscalYear + " - " + majlisName,
      htmlBody: htmlBody,
      name: SENDER_NAME
    });
    
    Logger.log("✅ Email sent to " + majlisName);
    ss.toast("Email sent to " + majlisName, "✅ Success", 3);
    
  } catch (error) {
    Logger.log("❌ Error: " + error.toString());
    SpreadsheetApp.getActiveSpreadsheet().toast("Error: " + error.toString(), "❌ Failed", 5);
  }
}

// ========================================
// BUILD MAJLIS EMAIL HTML
// ========================================
function buildMajlisEmailHTML(majlisName, khuddamData, atfalData, fiscalYear) {
  // TJ Column indexes: Anzahl=3, Nicht-Zahler=4, Budget=5, Bezahlt=18, Rest=19, Prozent=20
  
  // Khuddam data
  const kAnzahl = khuddamData ? khuddamData[3] : 0;
  const kNichtZahler = khuddamData ? khuddamData[4] : 0;
  const kBudget = khuddamData ? khuddamData[5] : 0;
  const kBezahlt = khuddamData ? khuddamData[18] : 0;
  const kRest = khuddamData ? khuddamData[19] : 0;
  const kProzent = khuddamData ? (khuddamData[20] * 100).toFixed(0) : 0;
  
  // Atfal data
  const aAnzahl = atfalData ? atfalData[3] : 0;
  const aNichtZahler = atfalData ? atfalData[4] : 0;
  const aBudget = atfalData ? atfalData[5] : 0;
  const aBezahlt = atfalData ? atfalData[18] : 0;
  const aRest = atfalData ? atfalData[19] : 0;
  const aProzent = atfalData ? (atfalData[20] * 100).toFixed(0) : 0;
  
  // Totals
  const tAnzahl = kAnzahl + aAnzahl;
  const tNichtZahler = kNichtZahler + aNichtZahler;
  const tBudget = kBudget + aBudget;
  const tBezahlt = kBezahlt + aBezahlt;
  const tRest = kRest + aRest;
  const tProzent = tBudget > 0 ? ((tBezahlt / tBudget) * 100).toFixed(0) : 0;
  
  return `
<!DOCTYPE html>
<html>
<head>
  <style>
    body { font-family: Arial, sans-serif; }
    table { border-collapse: collapse; }
    .header-table { width: 100%; border: none; margin-bottom: 20px; }
    .header-table td { border: none; padding: 10px; }
    .data-table { width: 100%; border: 1px solid #333; }
    .data-table th, .data-table td { border: 1px solid #333; padding: 8px; text-align: center; }
    .data-table th { background-color: #4285F4; color: white; font-weight: bold; }
    .title-row { background-color: #E8F0FE; font-weight: bold; }
  </style>
</head>
<body>
  <table class="header-table">
    <tr>
      <td width="150"><img src="${LOGO_URL}" alt="Logo" width="120" /></td>
      <td><strong>Majlis Khuddam-ul-Ahmadiyya Deutschland<br>Bundesverband<br>Chanda Department</strong></td>
    </tr>
  </table>
  
  <h2>Assalam-o-Alaikum</h2>
  
  <table class="data-table">
    <tr class="title-row">
      <td colspan="4"><h3>Chanda Summary Jahr ${fiscalYear}</h3></td>
    </tr>
    <tr class="title-row">
      <td colspan="2"><strong>Majlis</strong></td>
      <td colspan="2"><strong>${majlisName}</strong></td>
    </tr>
    <tr>
      <th>Feld</th>
      <th>Khuddam</th>
      <th>Atfal</th>
      <th>Total</th>
    </tr>
    <tr>
      <td><strong>Anzahl</strong></td>
      <td>${kAnzahl}</td>
      <td>${aAnzahl}</td>
      <td>${tAnzahl}</td>
    </tr>
    <tr>
      <td><strong>Nicht-Zahler</strong></td>
      <td>${kNichtZahler}</td>
      <td>${aNichtZahler}</td>
      <td>${tNichtZahler}</td>
    </tr>
    <tr>
      <td><strong>Budget</strong></td>
      <td>${formatCurrency(kBudget)}</td>
      <td>${formatCurrency(aBudget)}</td>
      <td>${formatCurrency(tBudget)}</td>
    </tr>
    <tr>
      <td><strong>Bezahlt</strong></td>
      <td>${formatCurrency(kBezahlt)}</td>
      <td>${formatCurrency(aBezahlt)}</td>
      <td>${formatCurrency(tBezahlt)}</td>
    </tr>
    <tr>
      <td><strong>Rest</strong></td>
      <td>${formatCurrency(kRest)}</td>
      <td>${formatCurrency(aRest)}</td>
      <td>${formatCurrency(tRest)}</td>
    </tr>
    <tr>
      <td><strong>%</strong></td>
      <td>${kProzent}%</td>
      <td>${aProzent}%</td>
      <td>${tProzent}%</td>
    </tr>
  </table>
  
  <p style="margin-top: 30px;">${EMAIL_SIGNATURE}</p>
</body>
</html>
  `;
}

// ========================================
// SEND REGION EMAIL
// ========================================
function sendRegionEmail(row) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const emailSheet = ss.getSheetByName(REGION_EMAIL_SHEET);
    const dataSheet = ss.getSheetByName(MAJALIS_DATA_SHEET);
    
    if (!emailSheet || !dataSheet) {
      throw new Error("Required sheets not found!");
    }
    
    // Get Region name and email
    const regionName = emailSheet.getRange(row, 1).getValue();
    const emailAddress = emailSheet.getRange(row, 2).getValue();
    
    if (!emailAddress || emailAddress === "") {
      throw new Error("No email address found for " + regionName);
    }
    
    Logger.log("Sending email to: " + regionName + " (" + emailAddress + ")");
    
    // Get data from Majalis_Data
    const data = dataSheet.getDataRange().getValues();
    const rows = data.slice(1);
    
    // Find all rows for this Region
    const regionRows = rows.filter(row => row[0] === regionName); // Col A = Region
    
    if (regionRows.length === 0) {
      throw new Error("No data found for " + regionName);
    }
    
    // Group by Majlis
    const majlisMap = {};
    for (let row of regionRows) {
      const majlis = row[1];
      const tanziem = row[2];
      
      if (!majlisMap[majlis]) {
        majlisMap[majlis] = { khuddam: null, atfal: null };
      }
      
      if (tanziem.toLowerCase().includes("khuddam")) {
        majlisMap[majlis].khuddam = row;
      } else if (tanziem.toLowerCase().includes("atfal")) {
        majlisMap[majlis].atfal = row;
      }
    }
    
    // Build and send email
    const fiscalYear = getFiscalYear();
    const htmlBody = buildRegionEmailHTML(regionName, majlisMap, fiscalYear);
    
    MailApp.sendEmail({
      to: emailAddress,
      subject: "Chanda Summary " + fiscalYear + " - " + regionName,
      htmlBody: htmlBody,
      name: SENDER_NAME
    });
    
    Logger.log("✅ Email sent to " + regionName);
    ss.toast("Email sent to " + regionName, "✅ Success", 3);
    
  } catch (error) {
    Logger.log("❌ Error: " + error.toString());
    SpreadsheetApp.getActiveSpreadsheet().toast("Error: " + error.toString(), "❌ Failed", 5);
  }
}

// ========================================
// BUILD REGION EMAIL HTML
// ========================================
function buildRegionEmailHTML(regionName, majlisMap, fiscalYear) {
  let tableRows = "";
  
  for (let majlis in majlisMap) {
    const kData = majlisMap[majlis].khuddam;
    const aData = majlisMap[majlis].atfal;
    
    // TJ: Anzahl=3, Nicht-Zahler=4, Budget=5, Bezahlt=18, Rest=19, Prozent=20
    const kAnzahl = kData ? kData[3] : 0;
    const kNichtZahler = kData ? kData[4] : 0;
    const kBudget = kData ? kData[5] : 0;
    const kBezahlt = kData ? kData[18] : 0;
    const kRest = kData ? kData[19] : 0;
    const kProzent = kData ? (kData[20] * 100).toFixed(0) : 0;
    
    const aAnzahl = aData ? aData[3] : 0;
    const aNichtZahler = aData ? aData[4] : 0;
    const aBudget = aData ? aData[5] : 0;
    const aBezahlt = aData ? aData[18] : 0;
    const aRest = aData ? aData[19] : 0;
    const aProzent = aData ? (aData[20] * 100).toFixed(0) : 0;
    
    tableRows += `
    <tr>
      <td><strong>${majlis}</strong></td>
      <td>${kAnzahl}</td>
      <td>${kNichtZahler}</td>
      <td>${formatCurrency(kBudget)}</td>
      <td>${formatCurrency(kBezahlt)}</td>
      <td>${formatCurrency(kRest)}</td>
      <td>${kProzent}%</td>
      <td>${aAnzahl}</td>
      <td>${aNichtZahler}</td>
      <td>${formatCurrency(aBudget)}</td>
      <td>${formatCurrency(aBezahlt)}</td>
      <td>${formatCurrency(aRest)}</td>
      <td>${aProzent}%</td>
    </tr>
    `;
  }
  
  return `
<!DOCTYPE html>
<html>
<head>
  <style>
    body { font-family: Arial, sans-serif; }
    table { border-collapse: collapse; }
    .header-table { width: 100%; border: none; margin-bottom: 20px; }
    .header-table td { border: none; padding: 10px; }
    .data-table { width: 100%; border: 1px solid #333; font-size: 11px; }
    .data-table th, .data-table td { border: 1px solid #333; padding: 6px; text-align: center; }
    .data-table th { background-color: #4285F4; color: white; font-weight: bold; }
    .title-row { background-color: #E8F0FE; font-weight: bold; }
  </style>
</head>
<body>
  <table class="header-table">
    <tr>
      <td width="150"><img src="${LOGO_URL}" alt="Logo" width="120" /></td>
      <td><strong>Majlis Khuddam-ul-Ahmadiyya Deutschland<br>Bundesverband<br>Chanda Department</strong></td>
    </tr>
  </table>
  
  <h2>Assalam-o-Alaikum</h2>
  
  <table class="data-table">
    <tr class="title-row">
      <td colspan="13"><h3>Chanda Summary Jahr ${fiscalYear}</h3></td>
    </tr>
    <tr class="title-row">
      <td colspan="13"><strong>Region: ${regionName}</strong></td>
    </tr>
    <tr>
      <th rowspan="2">Majlis</th>
      <th colspan="6">Khuddam</th>
      <th colspan="6">Atfal</th>
    </tr>
    <tr>
      <th>Tajneed</th>
      <th>Nicht-Zahler</th>
      <th>Budget</th>
      <th>Bezahlt</th>
      <th>Rest</th>
      <th>%</th>
      <th>Tajneed</th>
      <th>Nicht-Zahler</th>
      <th>Budget</th>
      <th>Bezahlt</th>
      <th>Rest</th>
      <th>%</th>
    </tr>
    ${tableRows}
  </table>
  
  <p style="margin-top: 30px;">${EMAIL_SIGNATURE}</p>
</body>
</html>
  `;
}

// ========================================
// HELPER: FORMAT CURRENCY
// ========================================
function formatCurrency(value) {
  if (typeof value !== "number") return "0.00";
  return value.toFixed(2).replace(/\B(?=(\d{3})+(?!\d))/g, ",");
}

// ========================================
// BULK EMAIL FUNCTIONS
// ========================================
function sendAllMajlisEmails() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(MAJLIS_EMAIL_SHEET);
  const data = sheet.getDataRange().getValues();
  
  let count = 0;
  for (let i = 2; i <= data.length; i++) {
    const email = sheet.getRange(i, 2).getValue();
    if (email && email !== "") {
      sendMajlisEmail(i);
      count++;
      Utilities.sleep(1000); // Wait 1 second between emails
    }
  }
  
  ss.toast("Sent " + count + " emails", "✅ Complete", 5);
}

function sendAllRegionEmails() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(REGION_EMAIL_SHEET);
  const data = sheet.getDataRange().getValues();
  
  let count = 0;
  for (let i = 2; i <= data.length; i++) {
    const email = sheet.getRange(i, 2).getValue();
    if (email && email !== "") {
      sendRegionEmail(i);
      count++;
      Utilities.sleep(1000);
    }
  }
  
  ss.toast("Sent " + count + " emails", "✅ Complete", 5);
}

// ========================================
// MENU
// ========================================
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('📧 Email Actions')
    .addItem('📨 Send All Majlis Emails', 'sendAllMajlisEmails')
    .addItem('📨 Send All Region Emails', 'sendAllRegionEmails')
    .addToUi();
}
