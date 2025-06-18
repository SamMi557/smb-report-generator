
// === CONFIGURATION ===
// Update with your actual Google Docs template and destination folder IDs
const TEMPLATE_ID = 'replace with template document id in google workspace';
const DESTINATION_FOLDER_ID = 'replace with destination folder-id in google workspace';

// === MENU SETUP ===
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('Generate Reports')
    .addItem('By Date Range', 'showDateRangePrompt')
    .addToUi();
}

// === PROMPT USER FOR DATE RANGE ===
function showDateRangePrompt() {
  const ui = SpreadsheetApp.getUi();

  const startPrompt = ui.prompt('Generate Reports', 'Enter the START date (YYYY-MM-DD):', ui.ButtonSet.OK_CANCEL);
  if (startPrompt.getSelectedButton() !== ui.Button.OK) return;

  const endPrompt = ui.prompt('Generate Reports', 'Enter the END date (YYYY-MM-DD):', ui.ButtonSet.OK_CANCEL);
  if (endPrompt.getSelectedButton() !== ui.Button.OK) return;

  const startDate = new Date(startPrompt.getResponseText());
  const endDate = new Date(endPrompt.getResponseText());

  if (isNaN(startDate) || isNaN(endDate)) {
    ui.alert('Please enter valid dates in the format YYYY-MM-DD.');
    return;
  }

  generateReportsInDateRange(startDate, endDate);
}

// === GENERATE REPORTS FOR ROWS IN DATE RANGE ===
function generateReportsInDateRange(startDate, endDate) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  const dateColIndex = headers.indexOf("Date");

  if (dateColIndex === -1) {
    SpreadsheetApp.getUi().alert('The column "Date" was not found.');
    return;
  }

  for (let i = 1; i < data.length; i++) {
    const rowDate = new Date(data[i][dateColIndex]);
    if (rowDate >= startDate && rowDate <= endDate) {
      const namedValues = {};
      headers.forEach((header, colIdx) => {
        namedValues[header] = [data[i][colIdx]?.toString() || ""];
      });
      onFormSubmit({ namedValues });
    }
  }
}

// === HANDLE FORM SUBMISSION OR TEST EVENT ===
function onFormSubmit(e) {
  if (!e || !e.namedValues) {
    throw new Error("Missing form submission data.");
  }

  const responses = e.namedValues;
  const cleanedResponses = cleanResponseKeys(responses);

  const company = cleanedResponses["Company Name"] || "UnknownCompany";
  const email = cleanedResponses["Email"] || "no-email@example.com";
  const date = new Date().toISOString().slice(0, 10);
  const fileName = `${company} - ${email} - ${date}`;

  const template = DriveApp.getFileById(TEMPLATE_ID);
  const folder = DriveApp.getFolderById(DESTINATION_FOLDER_ID);
  const docCopy = template.makeCopy(fileName, folder);
  const doc = DocumentApp.openById(docCopy.getId());
  const body = doc.getBody();

  for (const key in cleanedResponses) {
    const value = cleanedResponses[key] || "N/A";
    body.replaceText(`{{${key}}}`, value);
  }

  doc.saveAndClose();
  Logger.log(`Report created: ${fileName}`);
}

// === CLEAN FIELD NAME MAPPING ===
function cleanResponseKeys(namedValues) {
  const keyMap = {
    "Company Name ( Full legal business name)": "Company Name",
    "Email": "Email",
    "Primary Contact Name": "Primary Contact Name",
    "Role/Title": "Role",
    "Website": "Website",
    "Founder/CEO of company": "Founder Story",
    "Phone Number": "Phone",
    "LinkedIn Profile": "LinkedIn",
    "Company HQ Address": "Company Address",
    "Year Established": "Year Established",
    "Company Type": "Company Type",
    "Mission and Vision": "Mission",
    "Products and/or Services": "Products/Services",
    "Differentiators": "Differentiators",
    "Industry Sector": "Industry",
    "Stage of Business": "Stage",
    "Number of Full-Time Employees": "FT Employees",
    "Number of Part-Time Employees/Contractors": "PT Employees",
    "Monthly Revenue": "Revenue",
    "Profitability": "Profitability",
    "Key Team Member Bios": "Team Bios",
    "Areas needing help": "Help Areas",
    "Main Challenge": "Main Challenge",
    "What keeps the CEO up at night?": "CEO Concerns",
    "Top 3 Business Goals": "Goals",
    "Worked with consultants before?": "Worked with Consultants",
    "Supply Chain Model": "Supply Chain Model",
    "Supply Chain Description": "Supply Chain Details",
    "Sales Channels": "Sales Channels",
    "Customer Segments": "Customer Segments",
    "Please describe you customer persona & demographics.": "Customer Persona",
    "Top Markets": "Top Markets",
    "Competitors": "Competitors",
    "Technology Stack": "Technology Stack",
    "Last Fiscal Year Revenue (USD)": "Last Year Revenue",
    "Last Year Cost of Goods Sold (USD)": "Last Year COGS",
    "Last Year Operating Expenses (USD)": "Last Year OpEx",
    "Last Year Net Income (USD)": "Last Year Net Income",
    "Last Year Cash on Hand (USD)": "Last Year Cash",
    "Current YTD  Revenue (USD)": "YTD Revenue",
    "Current YTD  Cost of Goods Sold (USD)": "YTD COGS",
    "Current YTD Operating Expenses (USD)": "YTD OpEx",
    "Current YTD Net Income (USD)": "YTD Net Income",
    "Current YTD Cash on Hand (USD)": "YTD Cash",
    "Own IP?": "Own IP",
    "IP Details": "IP Details",
    "Legal Concerns?": "Legal Issues",
    "Legal Details": "Legal Details",
    "Company Impact Areas": "Impact Areas",
    "Aligned with SDGs?": "SDG Alignment",
    "If Yes, List SDGs": "SDG List",
    "Preferred Project Timeline": "Project Timeline",
    "Level of Involvement": "Involvement Level",
    "Are you open to let student(s) work on your problem statements?": "Student Engagement",
    "What does success look like?": "Success Definition",
    "Consent & Use": "Consent & Use",
    "Digital Signature": "Digital Signature",
    "Date": "Date"
  };

  const cleaned = {};
  for (let longKey in namedValues) {
    const shortKey = keyMap[longKey] || longKey;
    cleaned[shortKey] = namedValues[longKey].join(", ");
  }
  return cleaned;
}
