/**
 * DOST-3 Unified Form — Google Apps Script Backend
 * =====================================================
 * SETUP INSTRUCTIONS:
 * 1. Go to script.google.com and create a new project
 * 2. Paste this entire file as Code.gs
 * 3. Replace SPREADSHEET_ID below with your Google Sheet ID
 * 4. Deploy → New Deployment → Web App
 *    - Execute as: Me
 *    - Who has access: Anyone
 * 5. Copy the deployment URL and paste it into index.html and edit.html
 *    where it says: 'https://script.google.com/macros/s/YOUR_DEPLOYMENT_ID_HERE/exec'
 * 6. IMPORTANT: Run the "authorizeScript" function ONCE manually to grant Gmail permission!
 */

// ============================================================
// CONFIGURATION — EDIT THIS
// ============================================================
const SPREADSHEET_ID = '1k3sFNzBv1SiPohhSkuEdf-Tr7XAMNMsUwdyXg_vKdpM';
const SUBMISSIONS_SHEET = 'Submissions';
const UPDATES_SHEET = 'Updates';
const ADMIN_EMAIL = 'abenedicto431@gmail.com';

// ============================================================
// COLUMN HEADERS FOR SUBMISSIONS SHEET
// ============================================================
const SUBMISSION_HEADERS = [
  'Reference Number',
  'Submission Date',
  'Programs Applied',
  'Status',
  'Firm / Farm Name',
  'Province',
  'Owner Name',
  'Contact Person',
  'Owner Sex',
  'Owner Age',
  'Owner Position',
  'Contact Sex',
  'PWD',
  'Senior Citizen',
  'Complete Address',
  'Contact Number',
  'Email Address',
  'Birth Date',
  'Product/Service 1',
  'Product/Service 2',
  'Product/Service 3',
  'Dedicated Production Facility',
  'Facility Location',
  'SETUP iFund Beneficiary',
  'iFund Year Granted',
  'iFund Amount',
  'iFund Equipment',
  'Firm Background',
  'Reasons for TA',
  'Previously Consulted',
  'No Consult Reason',
  'Year Established',
  'Initial Capital',
  'Company Reg No',
  'Year Registered',
  'Organization Type',
  'Capital Classification',
  'Employment Classification',
  'Direct Workers - Production',
  'Direct Workers - Non-Production',
  'Indirect Workers',
  'Total Workers',
  'Annual Production Volume',
  'Estimated Value PHP',
  'Market - Direct Export',
  'Market - Sub-contractor',
  'Market - Potential Export',
  'Market - Local',
  'Market - Foreign',
  'Existing Foreign Market',
  'Existing Local Market',
  'Target Additional Market',
  'Business Activity',
  'Business Activity Others',
  'APP Products/Crops/Livestock',
  'APP Farm Plan',
  'MPP Products/Services',
  'MPP Firm Plan',
  'MPP Org Chart Available',
  'MPP Strategies and Problems',
  'EMP Electricity Consumption',
  'EMP Fuel Consumption',
  'EMP Combined Consumption',
  'Accomplished By',
  'Accomplished Date',
  'Edit Link',
  'Staff Status',
  'Staff Assigned',
  'Staff Remarks',
  'Staff Assessment Date',
  'Staff Decline Reason',
  'Assisted By',
  'Noted By',
  'Noted Date',
  'Last Updated',
];

// ============================================================
// COLUMN HEADERS FOR UPDATES SHEET (audit log)
// ============================================================
const UPDATE_HEADERS = [
  'Timestamp',
  'Reference Number',
  'Action',
  'Updated By',
  'Old Status',
  'New Status',
  'Notes',
];

// ============================================================
// RUN THIS FUNCTION ONCE MANUALLY TO AUTHORIZE GMAIL
// ============================================================
function authorizeScript() {
  // This forces Gmail and Spreadsheet authorization
  try {
    GmailApp.sendEmail(
      ADMIN_EMAIL,
      '[DOST-3] Authorization Test',
      'This is a test email to authorize the DOST-3 form script. Gmail is now authorized!\n\nYou can delete this email.'
    );
    SpreadsheetApp.openById(SPREADSHEET_ID);
    Logger.log('Authorization successful! Gmail and Sheets are now authorized.');
  } catch (e) {
    Logger.log('Authorization error: ' + e.toString());
  }
}

// ============================================================
// MAIN ENTRY POINT
// ============================================================
function doPost(e) {
  try {
    let data;
    try {
      data = JSON.parse(e.postData.contents);
    } catch (parseErr) {
      return ContentService.createTextOutput(
        JSON.stringify({ success: false, error: 'Invalid JSON payload' })
      ).setMimeType(ContentService.MimeType.JSON);
    }

    const action = data.action || 'submit';

    if (action === 'update') {
      return handleUpdate(data);
    } else {
      return handleSubmit(data);
    }
  } catch (err) {
    Logger.log('doPost error: ' + err.toString());
    return ContentService.createTextOutput(
      JSON.stringify({ success: false, error: err.toString() })
    ).setMimeType(ContentService.MimeType.JSON);
  }
}

function doGet(e) {
  try {
    const ref = e.parameter.ref;
    if (!ref) {
      return ContentService.createTextOutput(
        JSON.stringify({ success: false, error: 'No ref provided' })
      ).setMimeType(ContentService.MimeType.JSON);
    }

    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = ss.getSheetByName(SUBMISSIONS_SHEET);
    if (!sheet) {
      return ContentService.createTextOutput(
        JSON.stringify({ success: false, error: 'Submissions sheet not found' })
      ).setMimeType(ContentService.MimeType.JSON);
    }

    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    const data = sheet.getDataRange().getValues();
    const refIdx = headers.indexOf('Reference Number');

    for (let i = 1; i < data.length; i++) {
      if (data[i][refIdx] === ref) {
        const row = {};
        headers.forEach((h, j) => { row[h] = data[i][j]; });
        return ContentService.createTextOutput(
          JSON.stringify({ success: true, data: row })
        ).setMimeType(ContentService.MimeType.JSON);
      }
    }

    return ContentService.createTextOutput(
      JSON.stringify({ success: false, error: 'Submission not found' })
    ).setMimeType(ContentService.MimeType.JSON);

  } catch (err) {
    return ContentService.createTextOutput(
      JSON.stringify({ success: false, error: err.toString() })
    ).setMimeType(ContentService.MimeType.JSON);
  }
}

// ============================================================
// HANDLE NEW SUBMISSION
// ============================================================
function handleSubmit(data) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  let sheet = ss.getSheetByName(SUBMISSIONS_SHEET);

  if (!sheet) {
    sheet = ss.insertSheet(SUBMISSIONS_SHEET);
    sheet.appendRow(SUBMISSION_HEADERS);
    sheet.setFrozenRows(1);
    formatHeaderRow(sheet);
  } else if (sheet.getLastRow() === 0) {
    sheet.appendRow(SUBMISSION_HEADERS);
    sheet.setFrozenRows(1);
    formatHeaderRow(sheet);
  }

  const row = [
    data.referenceNumber || '',
    data.submissionDate || new Date().toISOString(),
    data.programs || '',
    'Pending',
    data.firmName || data.appFarmName || '',
    data.appProvince || data.firmProvince || data.fsProvince || '',
    data.ownerName || data.firmOwnerName || data.fsOwnerName || '',
    data.contactPerson || data.firmContactPerson || data.fsContactPerson || '',
    data.ownerSex || data.firmOwnerSex || '',
    data.ownerAge || data.firmOwnerAge || '',
    data.ownerPosition || data.firmOwnerPosition || '',
    data.contactSex || data.firmContactSex || '',
    data.ownerPwd || data.firmOwnerPwd || '',
    data.ownerSenior || data.firmOwnerSenior || '',
    data.appFarmAddress || data.firmAddress || data.fsFirmAddress || '',
    data.firmContactNum || data.empContactNum || data.fsContactNum || '',
    data.ownerEmail || data.firmEmail || data.fsEmail || '',
    data.fsOwnerBirthdate || '',
    data.fsProd1 || '',
    data.fsProd2 || '',
    data.fsProd3 || '',
    '',
    '',
    '',
    '',
    '',
    '',
    data.appBackground || data.mppBackground || data.empBackground || data.fsBackground || '',
    data.reasonsForTA || '',
    data.consultAns || data.mppConsultAns || data.empConsultAns || '',
    data.consultNoReason || data.mppConsultNoReason || data.empConsultNoReason || '',
    data.appYearEst || data.mppPaYear || data.empPaYear || '',
    data.mppPaCapital || data.empPaCapital || '',
    data.mppPaRegnum || data.empPaRegnum || '',
    data.mppPaRegyear || data.empPaRegyear || '',
    data.mppOrg || data.empOrg || data.fsOrg || '',
    data.mppCap || data.empCap || data.fsCap || '',
    data.mppEmp || data.empEmpcl || data.fsEmpcl || '',
    data.mppEmpProd || data.empEmpProd || '',
    data.mppEmpNonprod || data.empEmpNonprod || '',
    data.mppEmpIndirect || data.empEmpIndirect || '',
    data.mppEmpTotal || data.empEmpTotal || '',
    data.mppProdVol || data.empProdVol || '',
    data.mppProdVal || data.empProdVal || '',
    '',
    '',
    '',
    data.fsMktLocal || data.mppMktLocal || data.empMktLocal || '',
    data.fsMktForeign || data.mppMktForeign || data.empMktForeign || '',
    data.mppMktForeign || data.empMktForeign || '',
    data.mppMktLocal || data.empMktLocal || '',
    data.fsMktTarget || data.mppMktTarget || data.empMktTarget || '',
    data.mppConcerns || data.empConcerns || '',
    data.mppBizOthers || data.empBizOthers || '',
    data.appC1 || '',
    data.appFarmRows || '',
    data.mppProducts || data.empProducts || '',
    data.mppPlan || '',
    data.mppOrgchart || '',
    data.mppConcerns || '',
    data.empElecTotalKwh || '',
    data.empLpgTotalKg || '',
    data.empGrandTotal || '',
    data.accomplishedBy || data.ownerName || '',
    data.accomplishedDate || '',
    data.editLink || '',
    '',
    '',
    '',
    '',
    '',
    '',
    '',
    '',
    new Date().toISOString(),
  ];

  sheet.appendRow(row);

  const lastRow = sheet.getLastRow();
  sheet.getRange(lastRow, 1, 1, sheet.getLastColumn())
    .setBackground('#e8f5e9');

  // Send notification email
  try {
    sendNotificationEmail(data);
  } catch (emailErr) {
    Logger.log('Email error: ' + emailErr.toString());
  }

  return ContentService.createTextOutput(
    JSON.stringify({
      success: true,
      referenceNumber: data.referenceNumber,
      message: 'Submission recorded successfully'
    })
  ).setMimeType(ContentService.MimeType.JSON);
}

// ============================================================
// HANDLE STAFF UPDATE
// ============================================================
function handleUpdate(data) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName(SUBMISSIONS_SHEET);

  if (!sheet) {
    return ContentService.createTextOutput(
      JSON.stringify({ success: false, error: 'Submissions sheet not found' })
    ).setMimeType(ContentService.MimeType.JSON);
  }

  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const allData = sheet.getDataRange().getValues();
  const refIdx = headers.indexOf('Reference Number');

  for (let i = 1; i < allData.length; i++) {
    if (allData[i][refIdx] === data.ref) {
      const fieldUpdates = {
        'Staff Status': data.status || '',
        'Staff Assigned': data.staffAssigned || '',
        'Staff Remarks': data.staffRemarks || '',
        'Staff Assessment Date': data.staffAssessmentDate || '',
        'Staff Decline Reason': data.staffDeclineReason || '',
        'Assisted By': data.sigAssisted || '',
        'Noted By': data.sigNoted || '',
        'Noted Date': data.sigNotedDate || '',
        'Last Updated': new Date().toISOString(),
        'Status': data.status || '',
        'Firm / Farm Name': data.firmName || allData[i][headers.indexOf('Firm / Farm Name')],
        'Owner Name': data.ownerName || allData[i][headers.indexOf('Owner Name')],
      };

      Object.entries(fieldUpdates).forEach(([header, value]) => {
        const colIdx = headers.indexOf(header);
        if (colIdx >= 0 && value !== undefined) {
          sheet.getRange(i + 1, colIdx + 1).setValue(value);
        }
      });

      const statusColors = {
        pending: '#fff8e1',
        reviewed: '#e8f5e9',
        approved: '#e3f2fd',
        declined: '#fce4ec',
      };
      const color = statusColors[data.status] || '#ffffff';
      sheet.getRange(i + 1, 1, 1, sheet.getLastColumn()).setBackground(color);

      logUpdate(ss, data);

      return ContentService.createTextOutput(
        JSON.stringify({ success: true, message: 'Submission updated' })
      ).setMimeType(ContentService.MimeType.JSON);
    }
  }

  return ContentService.createTextOutput(
    JSON.stringify({ success: false, error: 'Reference number not found: ' + data.ref })
  ).setMimeType(ContentService.MimeType.JSON);
}

// ============================================================
// LOG STAFF UPDATES
// ============================================================
function logUpdate(ss, data) {
  let updateSheet = ss.getSheetByName(UPDATES_SHEET);
  if (!updateSheet) {
    updateSheet = ss.insertSheet(UPDATES_SHEET);
    updateSheet.appendRow(UPDATE_HEADERS);
    updateSheet.setFrozenRows(1);
  }
  updateSheet.appendRow([
    new Date().toISOString(),
    data.ref || '',
    'update',
    data.staffAssigned || 'Unknown Staff',
    '',
    data.status || '',
    data.staffRemarks || '',
  ]);
}

// ============================================================
// SEND EMAIL NOTIFICATION
// ============================================================
function sendNotificationEmail(data) {
  const ref = data.referenceNumber || 'N/A';
  const programs = data.programs || 'N/A';
  const firmName = data.firmName || data.appFarmName || data.fsFirmName || 'N/A';
  const province = data.appProvince || data.firmProvince || data.fsProvince || 'N/A';
  const ownerName = data.ownerName || data.firmOwnerName || data.fsOwnerName || 'N/A';
  const contactNum = data.firmContactNum || data.empContactNum || data.fsContactNum || 'N/A';
  const email = data.ownerEmail || data.firmEmail || data.fsEmail || 'N/A';
  const address = data.appFarmAddress || data.firmAddress || data.fsFirmAddress || 'N/A';
  const submittedDate = data.submissionDate ? new Date(data.submissionDate).toLocaleString('en-PH') : new Date().toLocaleString('en-PH');

  const subject = `[DOST-3] New Submission: ${ref} — ${firmName}`;

  const body = `
==============================================
  DOST-3 NEW APPLICATION SUBMISSION
==============================================

Reference Number : ${ref}
Programs Applied : ${programs}
Submission Date  : ${submittedDate}

----------------------------------------------
APPLICANT INFORMATION
----------------------------------------------
Firm / Farm Name : ${firmName}
Province         : ${province}
Owner Name       : ${ownerName}
Contact Number   : ${contactNum}
Email Address    : ${email}
Address          : ${address}

----------------------------------------------
PROGRAM-SPECIFIC DETAILS
----------------------------------------------
${programs.includes('APP') ? `
[APP - Agricultural Productivity Program]
Farm Name        : ${data.appFarmName || 'N/A'}
Farm Address     : ${data.appFarmAddress || 'N/A'}
Farm Area        : ${data.appArea || 'N/A'}
Year Established : ${data.appYearEst || 'N/A'}
No. of Workers   : ${data.appWorkers || 'N/A'}
Farm Background  : ${data.appBackground || 'N/A'}
` : ''}
${programs.includes('MPP') ? `
[MPP - Manufacturing Productivity Program]
Firm Name        : ${data.firmName || 'N/A'}
Firm Address     : ${data.firmAddress || 'N/A'}
Year Established : ${data.mppPaYear || 'N/A'}
Initial Capital  : ${data.mppPaCapital || 'N/A'}
Reg. Number      : ${data.mppPaRegnum || 'N/A'}
Organization     : ${data.mppOrg || 'N/A'}
Capital Class    : ${data.mppCap || 'N/A'}
Employment Class : ${data.mppEmp || 'N/A'}
Total Employees  : ${data.mppEmpTotal || 'N/A'}
Products/Services: ${data.mppProducts || 'N/A'}
Org Chart        : ${data.mppOrgchart || 'N/A'}
Concerns         : ${data.mppConcerns || 'N/A'}
` : ''}
${programs.includes('EMP') ? `
[EMP - Energy Management Program]
Firm Name        : ${data.empFirmName || data.firmName || 'N/A'}
Year Established : ${data.empPaYear || 'N/A'}
Total Employees  : ${data.empEmpTotal || 'N/A'}
Elec Total kWh   : ${data.empElecTotalKwh || 'N/A'}
Elec Total PHP   : ${data.empElecTotalPhp || 'N/A'}
LPG Total kg     : ${data.empLpgTotalKg || 'N/A'}
Diesel Total L   : ${data.empDieselTotalLiters || 'N/A'}
Grand Total PHP  : ${data.empGrandTotal || 'N/A'}
` : ''}
${programs.includes('FS') ? `
[FS - Food Safety Enrollment Form]
Firm Name        : ${data.fsFirmName || 'N/A'}
Firm Address     : ${data.fsFirmAddress || 'N/A'}
Owner Name       : ${data.fsOwnerName || 'N/A'}
Product 1        : ${data.fsProd1 || 'N/A'}
Product 2        : ${data.fsProd2 || 'N/A'}
Product 3        : ${data.fsProd3 || 'N/A'}
Organization     : ${data.fsOrg || 'N/A'}
` : ''}
----------------------------------------------
STAFF EDIT LINK (Authorized Personnel Only)
----------------------------------------------
${data.editLink || 'N/A'}

==============================================
This is an automated notification from the
DOST-3 Online Form System.
==============================================
  `.trim();

  GmailApp.sendEmail(ADMIN_EMAIL, subject, body);
  Logger.log('Email sent successfully to ' + ADMIN_EMAIL);
}

// ============================================================
// FORMAT HEADER ROW
// ============================================================
function formatHeaderRow(sheet) {
  const headerRange = sheet.getRange(1, 1, 1, SUBMISSION_HEADERS.length);
  headerRange
    .setBackground('#1a2d4f')
    .setFontColor('#ffffff')
    .setFontWeight('bold')
    .setFontSize(10);
  sheet.setFrozenRows(1);
  sheet.setColumnWidth(1, 160);
  sheet.setColumnWidth(2, 170);
  sheet.setColumnWidth(3, 120);
  sheet.setColumnWidth(4, 100);
  sheet.setColumnWidth(5, 200);
  sheet.setColumnWidth(6, 100);
}

// ============================================================
// TEST FUNCTION
// ============================================================
function testSubmit() {
  const testData = {
    referenceNumber: 'DOST3-TEST01',
    submissionDate: new Date().toISOString(),
    programs: 'MPP, EMP',
    firmName: 'Test Manufacturing Corp',
    firmProvince: 'Pampanga',
    ownerName: 'Juan dela Cruz',
    contactPerson: 'Maria Santos',
    ownerSex: 'Male',
    ownerAge: '45',
    ownerPosition: 'Owner',
    firmAddress: '123 Brgy. Test, Angeles City, Pampanga',
    firmContactNum: '09171234567',
    firmEmail: 'test@example.com',
    mppBackground: 'A test manufacturing firm.',
    mppPaYear: '2010',
    mppPaCapital: '5,000,000',
    mppPaRegnum: 'DTI-12345',
    mppPaRegyear: '2010',
    mppOrg: 'Single Proprietorship',
    mppCap: 'Small — P 1.5–15 M',
    mppEmp: 'Small (10–99)',
    mppEmpProd: '20',
    mppEmpNonprod: '5',
    mppEmpIndirect: '3',
    mppEmpTotal: '28',
    mppProducts: 'Metal fabricated parts',
    editLink: 'https://abenedicto431-glitch.github.io/PORMS/edit.html?ref=DOST3-TEST01',
  };

  handleSubmit(testData);
  Logger.log('Test submission completed. Check your Google Sheet and Gmail.');
}
