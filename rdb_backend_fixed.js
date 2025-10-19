/**
 * RDB Case Management Backend - Production v1.2
 * Matches intake-logic.js field names exactly
 * Populates: Cases, Clients, Attorneys, Employers, Claimants
 */

const SHEET_NAMES = ["Clients", "Attorneys", "Employers", "Claimants", "Cases", "Investigators", "Interviews", "Reports", "Updates", "Billing"];

function generateId_(sheet) {
  const name = sheet.getName();
  const prefix = name.substring(0, 2).toUpperCase();
  const data = sheet.getDataRange().getValues();
  const lastRow = data.length;
  if (lastRow < 2) return prefix + "-0001";
  const lastId = data[lastRow - 1][0];
  const num = parseInt(lastId.split("-")[1]) + 1;
  return prefix + "-" + num.toString().padStart(4, "0");
}

function doGet(e) {
  try {
    const sheetName = e.parameter.sheet;
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);

    if (!sheet) {
      return ContentService.createTextOutput(JSON.stringify({
        status: "error",
        message: "Sheet " + sheetName + " not found"
      })).setMimeType(ContentService.MimeType.JSON);
    }

    const data = sheet.getDataRange().getValues();
    const headers = data.shift();
    const objects = data.map(function(r) {
      const obj = {};
      headers.forEach(function(h, i) {
        obj[h] = r[i];
      });
      return obj;
    });

    return ContentService.createTextOutput(JSON.stringify({
      status: "success",
      data: objects
    })).setMimeType(ContentService.MimeType.JSON);

  } catch (err) {
    return ContentService.createTextOutput(JSON.stringify({
      status: "error",
      message: err.toString()
    })).setMimeType(ContentService.MimeType.JSON);
  }
}

function doPost(e) {
  try {
    const input = JSON.parse(e.postData.contents);
    let targetSheetName, rowData, recordId;

    // Legacy format: {sheet: "Cases", data: {...}}
    if (input.sheet && input.data) {
      targetSheetName = input.sheet;
      const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(targetSheetName);
      if (!sheet) throw new Error("Sheet " + targetSheetName + " not found");

      const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
      recordId = generateId_(sheet);
      rowData = headers.map(function(h) {
        return (h === headers[0] ? recordId : input.data[h] || "");
      });
      sheet.appendRow(rowData);

      // NEW FORMAT: {action: 'syncCase', caseData: {...}}
    } else if (input.action === 'syncCase' && input.caseData) {
      const cd = input.caseData;
      const ss = SpreadsheetApp.getActiveSpreadsheet();
      const ids = {};

      // 1. CREATE CLIENT RECORD (if client data exists)
      if (cd.clientName || cd.clientCompany) {
        const clientSheet = ss.getSheetByName("Clients");
        if (clientSheet) {
          const clientId = generateId_(clientSheet);
          ids.clientId = clientId;

          clientSheet.appendRow([
            clientId,                                    // ClientID
            cd.clientName || '',                        // Name (Adjuster Name)
          cd.clientCompany || '',                     // Company
          cd.clientTitle || '',                       // Title
          cd.clientPhone || '',                       // Phone
          cd.clientEmail || '',                       // Email
          cd.clientStreet || '',                      // Street
          cd.clientCity || '',                        // City
          cd.clientState || 'CA',                     // State
          cd.clientZip || '',                         // Zip
          new Date()                                   // DateCreated
          ]);
        }
      }

      // 2. CREATE ATTORNEY RECORD (if attorney data exists)
      if (cd.attorneyName || cd.attorneyCompany) {
        const attorneySheet = ss.getSheetByName("Attorneys");
        if (attorneySheet) {
          const attorneyId = generateId_(attorneySheet);
          ids.attorneyId = attorneyId;

          attorneySheet.appendRow([
            attorneyId,                                  // AttorneyID
            cd.attorneyName || '',                      // Name
            cd.attorneyCompany || '',                   // Company
            cd.attorneyTitle || 'Attorney',             // Title
            cd.attorneyPhone || '',                     // Phone
            cd.attorneyEmail || '',                     // Email
            cd.attorneyStreet || '',                    // Street
            cd.attorneyCity || '',                      // City
            cd.attorneyState || 'CA',                   // State
            cd.attorneyZip || '',                       // Zip
            new Date()                                   // DateCreated
          ]);
        }
      }

      // 3. CREATE EMPLOYER RECORD (if employer data exists)
      if (cd.employerName || cd.employerCompany) {
        const employerSheet = ss.getSheetByName("Employers");
        if (employerSheet) {
          const employerId = generateId_(employerSheet);
          ids.employerId = employerId;

          employerSheet.appendRow([
            employerId,                                  // EmployerID
            cd.employerName || '',                      // ContactName
            cd.employerCompany || '',                   // Company
            cd.employerTitle || '',                     // Title
            cd.employerPhone || '',                     // Phone
            cd.employerEmail || '',                     // Email
            cd.employerStreet || '',                    // Street
            cd.employerCity || '',                      // City
            cd.employerState || 'CA',                   // State
            cd.employerZip || '',                       // Zip
            new Date()                                   // DateCreated
          ]);
        }
      }

      // 4. CREATE CLAIMANT RECORD (if claimant data exists)
      if (cd.claimantName) {
        const claimantSheet = ss.getSheetByName("Claimants");
        if (claimantSheet) {
          const claimantId = generateId_(claimantSheet);
          ids.claimantId = claimantId;

          claimantSheet.appendRow([
            claimantId,                                  // ClaimantID
            cd.claimantName || '',                      // Name
            cd.claimantJobTitle || '',                  // JobTitle
            cd.claimantPhone || '',                     // Phone
            cd.claimantEmail || '',                     // Email
            cd.claimantSSN || '',                       // SSN
            cd.claimantDOB || '',                       // DOB
            cd.claimantStreet || '',                    // Street
            cd.claimantCity || '',                      // City
            cd.claimantState || 'CA',                   // State
            cd.claimantZip || '',                       // Zip
            new Date()                                   // DateCreated
          ]);
        }
      }

      // 5. CREATE CASE RECORD (always)
      const caseSheet = ss.getSheetByName("Cases");
      if (!caseSheet) throw new Error('Cases sheet not found');

      const caseId = generateId_(caseSheet);
      recordId = caseId;
      targetSheetName = "Cases";

      caseSheet.appendRow([
        caseId,                                          // CaseID
        cd.assignmentClaimNumber || '',                 // ClaimNumber
        ids.clientId || '',                             // ClientID (reference)
      ids.attorneyId || '',                           // AttorneyID (reference)
      ids.employerId || '',                           // EmployerID (reference)
      ids.claimantId || '',                           // ClaimantID (reference)
      cd.investigatorName || '',                      // InvestigatorAssigned
      cd.assignmentService || '',                     // Service
      cd.assignmentInjuryDate || '',                  // InjuryDate
      cd.assignmentAssignmentDate || new Date(),      // AssignmentDate
                          cd.assignmentDueDate || '',                     // DueDate
                          cd.assignmentDecisionDate || '',                // DecisionDate
                          cd.assignmentStatus || 'Open',                  // Status
                          cd.assignmentPipelineStage || '',               // PipelineStage
                          cd.caseClaimantStatement || '',                 // ClaimantStatement
                          cd.caseClaimantWorking || '',                   // ClaimantWorking
                          cd.caseMedicalRelease || '',                    // MedicalRelease
                          cd.caseRestrictions || '',                      // Restrictions
                          cd.caseInjuryDescription || '',                 // InjuryDescription
                          cd.caseAdjusterInstructions || '',              // AdjusterInstructions
                          new Date(),                                      // DateCreated
                          ''                                               // Notes
      ]);

    } else if (input.action === 'syncBilling' && input.billingData) {
      targetSheetName = "Billing";
      const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(targetSheetName);
      if (!sheet) throw new Error('Billing sheet not found');

      const bd = input.billingData;
      recordId = generateId_(sheet);

      rowData = [
        recordId,
        bd.claimNumber || '',
        bd.investigator || '',
        bd.billingAmount || '',
        bd.billingDate || new Date(),
        bd.invoiceNumber || '',
        bd.notes || ''
      ];
      sheet.appendRow(rowData);

    } else if (input.action === 'syncInterview' && input.interviewData) {
      targetSheetName = "Interviews";
      const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(targetSheetName);
      if (!sheet) throw new Error('Interviews sheet not found');

      const id = input.interviewData;
      recordId = generateId_(sheet);

      rowData = [
        recordId,
        id.claimNumber || '',
        id.investigationType || '',
        id.investigator || '',
        id.interviewDate || '',
        id.interviewTime || '',
        id.locationName || '',
        id.locationStreet || '',
        id.locationCity || '',
        id.locationState || 'CA',
        id.locationZip || '',
        id.notes || '',
        new Date()
      ];
      sheet.appendRow(rowData);

    } else if (input.action === 'syncReport' && input.reportData) {
      targetSheetName = "Reports";
      const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(targetSheetName);
      if (!sheet) throw new Error('Reports sheet not found');

      const rd = input.reportData;
      recordId = generateId_(sheet);

      rowData = [
        recordId,
        rd.claimNumber || '',
        rd.investigator || '',
        rd.submissionDate || new Date(),
        rd.lastInterviewDate || '',
        rd.notes || ''
      ];
      sheet.appendRow(rowData);

    } else if (input.action === 'syncUpdate' && input.updateData) {
      targetSheetName = "Updates";
      const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(targetSheetName);
      if (!sheet) throw new Error('Updates sheet not found');

      const ud = input.updateData;
      recordId = generateId_(sheet);

      rowData = [
        recordId,
        ud.claimNumber || '',
        ud.actionTaken || '',
        ud.nextSteps || '',
        ud.notes || '',
        ud.interviewDate || '',
        ud.statusBefore || '',
        ud.statusAfter || '',
        ud.updateDate || new Date(),
        ud.updatedBy || 'Conrad'
      ];
      sheet.appendRow(rowData);

    } else {
      throw new Error('Invalid payload format');
    }

    return ContentService.createTextOutput(JSON.stringify({
      status: "success",
      id: recordId,
      sheet: targetSheetName
    })).setMimeType(ContentService.MimeType.JSON);

  } catch (err) {
    return ContentService.createTextOutput(JSON.stringify({
      status: "error",
      message: err.toString()
    })).setMimeType(ContentService.MimeType.JSON);
  }
}

function getInvestigators() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Investigators");
  if (!sheet) return [];
  const data = sheet.getDataRange().getValues();
  data.shift();
  return data.filter(function(r) {
    return r[5] === 'Y';
  }).map(function(r) {
    return {
      InvestigatorID: r[0],
      Name: r[1],
      Phone: r[2],
      Email: r[3],
      CityOfResidence: r[4],
      Active: r[5]
    };
  });
}

function onDeploy() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let log = ss.getSheetByName("_Log");
  if (!log) {
    log = ss.insertSheet("_Log");
    log.hideSheet();
    log.appendRow(["Timestamp", "User", "Version"]);
  }
  log.appendRow([new Date(), Session.getActiveUser().getEmail() || "Unknown", "v1.2"]);
}
