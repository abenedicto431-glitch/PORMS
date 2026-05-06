/**
 * DOST-3 2025 Unified Form — Google Apps Script Backend
 * =====================================================
 * SETUP INSTRUCTIONS:
 * 1. Go to script.google.com and create a new project
 * 2. Paste this entire file as Code.gs
 * 3. Replace SPREADSHEET_ID below with your Google Sheet ID
 * 4. Deploy → New Deployment → Web App
 *    - Execute as: Me
 *    - Who has access: Anyone
 * 5. Copy the deployment URL and paste it into index.html and edit.html
 *    where it says: 'https://script.google.com/macros/s/https://script.google.com/macros/s/AKfycbxjop3hnUQV1rzl87d3hFQSyWs_Xy9RuNhT01yAe7YCkfH4sWiJTIjH20ahGkoSKL4/exec'
 *
 * SPREADSHEET SETUP:
 * - Create a Google Sheet with these sheet tabs (exact names):
 *   "Submissions" — main data table
 *   "Updates"     — staff edit log
 *
 * The script will auto-create headers on first run.
 */

// ============================================================
// CONFIGURATION — EDIT THIS
// ============================================================
const SPREADSHEET_ID = '1k3sFNzBv1SiPohhSkuEdf-Tr7XAMNMsUwdyXg_vKdpM'; // Replace with your Sheet ID
const SUBMISSIONS_SHEET = 'Submissions';
const UPDATES_SHEET = 'Updates';
const ADMIN_EMAIL = 'abenedicto431@gmail.com'; // Replace with your email

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
  // FS Products
  'Product/Service 1',
  'Product/Service 2',
  'Product/Service 3',
  // FS Facility
  'Dedicated Production Facility',
  'Facility Location',
  // SETUP iFund
  'SETUP iFund Beneficiary',
  'iFund Year Granted',
  'iFund Amount',
  'iFund Equipment',
  // Firm Profile
  'Firm Background',
  'Reasons for TA',
  // Previous Consultations
  'Previously Consulted',
  'No Consult Reason',
  // Pre-Assessment
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
  // APP
  'APP Products/Crops/Livestock',
  'APP Farm Plan',
  // MPP
  'MPP Products/Services',
  'MPP Firm Plan',
  'MPP Org Chart Available',
  'MPP Strategies and Problems',
  // EMP
  'EMP Electricity Consumption',
  'EMP Fuel Consumption',
  'EMP Combined Consumption',
  // Signatures
  'Accomplished By',
  'Accomplished Date',
  'Edit Link',
  // Staff fields (updated later)
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
// MAIN ENTRY POINT
// ============================================================
function doPost(e) {
  const corsHeaders = {
    'Access-Control-Allow-Origin': '*',
    'Access-Control-Allow-Methods': 'POST, OPTIONS',
    'Access-Control-Allow-Headers': 'Content-Type',
  };

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
      return handleUpdate(data, corsHeaders);
    } else {
      return handleSubmit(data, corsHeaders);
    }
  } catch (err) {
    Logger.log('doPost error: ' + err.toString());
    return ContentService.createTextOutput(
      JSON.stringify({ success: false, error: err.toString() })
    ).setMimeType(ContentService.MimeType.JSON);
  }
}

function doGet(e) {
  // Allow fetching a single submission by ref (for edit page pre-fill)
  try {
    const ref = e.parameter.ref;
    if (!ref) {
      return ContentService.createTextOutput(
        JSON.stringify({ success: false, error: 'No ref provided' })
      ).setMimeType(ContentService.MimeType.JSON);
    }

    const ss = SpreadsheetApp.openById('1k3sFNzBv1SiPohhSkuEdf-Tr7XAMNMsUwdyXg_vKdpM');
    const sheet = ss.getSheetByName(Submissions);
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
function handleSubmit(data, corsHeaders) {
  const ss = SpreadsheetApp.openById('1k3sFNzBv1SiPohhSkuEdf-Tr7XAMNMsUwdyXg_vKdpM');
  let sheet = ss.getSheetByName(Submissions);

  // Create sheet with headers if it doesn't exist
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

  // Build row values matching SUBMISSION_HEADERS
  const row = [
    data.referenceNumber || '',
    data.submissionDate || new Date().toISOString(),
    data.programs || '',
    'Pending',                         // Status — default
    data.firmName || '',
    data.province || '',
    data.ownerName || '',
    data.contactPerson || '',
    data.ownerSex || '',
    data.ownerAge || '',
    data.ownerPosition || '',
    data.contactSex || '',
    data.ownerPwd || '',
    data.ownerSenior || '',
    data.address || '',
    data.contactNumber || '',
    data.email || '',
    data.ownerBirthdate || '',
    data.fsProd1 || '',
    data.fsProd2 || '',
    data.fsProd3 || '',
    data.fsFacility || '',
    data.fsFacilityLoc || '',
    data.setupFund || '',
    data.setupYear || '',
    data.setupAmount || '',
    data.setupEquipment || '',
    data.firmBackground || '',
    data.reasonsForTA || '',
    data.prevConsult || '',
    data.consultNoReason || '',
    data.yearEstablished || '',
    data.initialCapital || '',
    data.regNumber || '',
    data.yearRegistered || '',
    data.orgType || '',
    data.capClass || '',
    data.empClass || '',
    data.empProduction || '',
    data.empNonprod || '',
    data.empIndirect || '',
    data.empTotal || '',
    data.prodVolume || '',
    data.prodValue || '',
    data.mktDirectExport || '',
    data.mktSubcontractor || '',
    data.mktPotentialExport || '',
    data.mktLocal || '',
    data.mktForeign || '',
    data.existingForeignMarket || '',
    data.existingLocalMarket || '',
    data.targetMarket || '',
    data.businessActivity || '',
    data.businessOthers || '',
    data.appProducts || '',
    data.appPlan || '',
    data.mppProducts || '',
    data.mppPlan || '',
    data.mppOrgChart || '',
    data.mppStrategies || '',
    data.empElectricity || '',
    data.empFuel || '',
    data.empCombined || '',
    data.accomplishedBy || '',
    data.accomplishedDate || '',
    data.editLink || '',
    '',  // Staff Status (filled later)
    '',  // Staff Assigned
    '',  // Staff Remarks
    '',  // Staff Assessment Date
    '',  // Staff Decline Reason
    '',  // Assisted By
    '',  // Noted By
    '',  // Noted Date
    new Date().toISOString(),  // Last Updated
  ];

  sheet.appendRow(row);

  // Color-code new row
  const lastRow = sheet.getLastRow();
  sheet.getRange(lastRow, 1, 1, sheet.getLastColumn())
    .setBackground('#e8f5e9');

  // Send notification email
  try {
    sendNotificationEmail(data);
  } catch (emailErr) {
    Logger.log('Email error: ' + emailErr);
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
function handleUpdate(data, corsHeaders) {
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
      // Update staff fields
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
        // Also allow updating core fields
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

      // Update row color based on status
      const statusColors = {
        pending: '#fff8e1',
        reviewed: '#e8f5e9',
        approved: '#e3f2fd',
        declined: '#fce4ec',
      };
      const color = statusColors[data.status] || '#ffffff';
      sheet.getRange(i + 1, 1, 1, sheet.getLastColumn()).setBackground(color);

      // Log to updates sheet
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
  const subject = `[DOST-3] New Submission: ${data.referenceNumber} — ${data.firmName || 'Unknown Firm'}`;
  const body = `
A new DOST-3 application has been submitted.

Reference Number: ${data.referenceNumber}
Programs: ${data.programs}
Firm/Farm: ${data.firmName}
Province: ${data.province}
Owner: ${data.ownerName}
Contact: ${data.contactNumber}
Email: ${data.email}
Submitted: ${new Date(data.submissionDate).toLocaleString('en-PH')}

Staff Edit Link:
${data.editLink}

---
This is an automated notification from the DOST-3 Online Form System.
  `.trim();

  GmailApp.sendEmail(ADMIN_EMAIL, subject, body);
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

  // Set column widths for key columns
  sheet.setColumnWidth(1, 160);  // Reference Number
  sheet.setColumnWidth(2, 170);  // Submission Date
  sheet.setColumnWidth(3, 120);  // Programs
  sheet.setColumnWidth(4, 100);  // Status
  sheet.setColumnWidth(5, 200);  // Firm Name
  sheet.setColumnWidth(6, 100);  // Province
}

// ============================================================
// TEST FUNCTION (run manually in Apps Script editor)
// ============================================================
function testSubmit() {
  const testData = {
    referenceNumber: 'DOST3-2025-TEST01',
    submissionDate: new Date().toISOString(),
    programs: 'MPP, EMP',
    firmName: 'Test Manufacturing Corp',
    province: 'Pampanga',
    ownerName: 'Juan dela Cruz',
    contactPerson: 'Maria Santos',
    ownerSex: 'Male',
    ownerAge: '45',
    ownerPosition: 'Owner',
    address: '123 Brgy. Test, Angeles City, Pampanga',
    contactNumber: '09171234567',
    email: 'test@example.com',
    firmBackground: 'A test manufacturing firm.',
    reasonsForTA: 'Seeking productivity improvement.',
    yearEstablished: '2010',
    initialCapital: '5,000,000',
    regNumber: 'DTI-12345',
    yearRegistered: '2010',
    accomplishedBy: 'Juan dela Cruz',
    accomplishedDate: new Date().toISOString().split('T')[0],
    editLink: 'https://yourgithubusername.github.io/DOST-FORMS/edit.html?ref=DOST3-2025-TEST01',
  };

  handleSubmit(testData, {});
  Logger.log('Test submission completed. Check your Google Sheet.');
}
