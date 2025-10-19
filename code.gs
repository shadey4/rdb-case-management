/**
 * RDB Case Management Backend - Production v1.2
 * Populates: Cases, Clients, Attorneys, Employers, Claimants
 * Adds duplicate prevention for Claim Numbers
 */

const SHEET_NAMES = [
    "Clients", "Attorneys", "Employers", "Claimants",
"Cases", "Investigators", "Interviews", "Reports", "Updates", "Billing"
];

const FIELD_MAPPINGS = {
    // Assignment fields
    assignmentClaimNumber: "Claim Number",
    assignmentService: "Service",
    assignmentInjuryDate: "Injury Date",
    assignmentAssignmentDate: "Assignment Date",
    assignmentDueDate: "Due Date",
    assignmentDecisionDate: "Decision Date",
    assignmentStatus: "Status",
    assignmentPipelineStage: "Pipeline Stage",

    // Client (Adjuster)
    clientName: "Adjuster Name",
    clientCompany: "Company",
    clientTitle: "Client Title",
    clientPhone: "Phone",
    clientEmail: "Email",
    clientStreet: "Street",
    clientCity: "City",
    clientState: "State",
    clientZip: "Zip",

    // Attorney
    attorneyName: "Attorney Name",
    attorneyCompany: "Attorney Firm",
    attorneyTitle: "Attorney Title",
    attorneyPhone: "Phone",
    attorneyEmail: "Email",
    attorneyStreet: "Street",
    attorneyCity: "City",
    attorneyState: "State",
    attorneyZip: "Zip",

    // Employer
    employerName: "Contact Name",
    employerCompany: "Company",
    employerTitle: "Employer Title",
    employerPhone: "Phone",
    employerEmail: "Email",
    employerStreet: "Street",
    employerCity: "City",
    employerState: "State",
    employerZip: "Zip",

    // Claimant
    claimantName: "Claimant Name",
    claimantJobTitle: "Job Title",
    claimantPhone: "Phone",
    claimantEmail: "Email",
    claimantSSN: "SSN",
    claimantDOB: "Date of Birth",
    claimantStreet: "Street",
    claimantCity: "City",
    claimantState: "State",
    claimantZip: "Zip",

    // Case Info
    caseClaimantStatement: "Claimant Statement",
        caseClaimantWorking: "Claimant Working",
            caseMedicalRelease: "Medical Release",
                caseRestrictions: "Restrictions",
                    caseInjuryDescription: "Injury Description",
                        caseAdjusterInstructions: "Adjuster Instructions",

                            investigatorName: "Investigator"
};

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

function translateFields_(data) {
    const result = {};
    for (const key in data) result[FIELD_MAPPINGS[key] || key] = data[key];
    return result;
}

function doGet(e) {
    try {
        const sheetName = e.parameter.sheet;
        const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
        if (!sheet) throw new Error("Sheet " + sheetName + " not found");
        const data = sheet.getDataRange().getValues();
        const headers = data.shift();
        const objects = data.map(r => {
            const obj = {};
            headers.forEach((h, i) => obj[h] = r[i]);
            return obj;
        });
        return ContentService.createTextOutput(JSON.stringify({ status: "success", data: objects }))
        .setMimeType(ContentService.MimeType.JSON);
    } catch (err) {
        return ContentService.createTextOutput(JSON.stringify({ status: "error", message: err.toString() }))
        .setMimeType(ContentService.MimeType.JSON);
    }
}

function doPost(e) {
    try {
        const input = JSON.parse(e.postData.contents);
        const ss = SpreadsheetApp.getActiveSpreadsheet();
        let recordId = null;

        if (input.action === 'syncCase' && input.caseData) {
            const cd = translateFields_(input.caseData);

            // === DUPLICATE PREVENTION ===
            const caseSheet = ss.getSheetByName("Cases");
            if (caseSheet) {
                const data = caseSheet.getDataRange().getValues();
                const claimNumberCol = 1; // Column B (Claim Number)
                for (let i = 1; i < data.length; i++) {
                    if (data[i][claimNumberCol] === cd["Claim Number"]) {
                        return ContentService.createTextOutput(JSON.stringify({
                            status: "success",
                            message: "Case already exists (duplicate prevented)",
                                                                              id: data[i][0]
                        })).setMimeType(ContentService.MimeType.JSON);
                    }
                }
            }

            const ids = {};
            // === CLIENT ===
            if (cd["Adjuster Name"] || cd["Company"]) {
                const s = ss.getSheetByName("Clients");
                if (s) {
                    const id = generateId_(s);
                    ids.clientId = id;
                    s.appendRow([id, cd["Adjuster Name"], cd["Company"], cd["Client Title"], cd["Phone"], cd["Email"],
                                cd["Street"], cd["City"], cd["State"] || 'CA', cd["Zip"], new Date()]);
                }
            }

            // === ATTORNEY ===
            if (cd["Attorney Name"] || cd["Attorney Firm"]) {
                const s = ss.getSheetByName("Attorneys");
                if (s) {
                    const id = generateId_(s);
                    ids.attorneyId = id;
                    s.appendRow([id, cd["Attorney Name"], cd["Attorney Firm"], cd["Attorney Title"] || 'Attorney',
                                cd["Phone"], cd["Email"], cd["Street"], cd["City"], cd["State"] || 'CA', cd["Zip"], new Date()]);
                }
            }

            // === EMPLOYER ===
            if (cd["Contact Name"] || cd["Company"]) {
                const s = ss.getSheetByName("Employers");
                if (s) {
                    const id = generateId_(s);
                    ids.employerId = id;
                    s.appendRow([id, cd["Contact Name"], cd["Company"], cd["Employer Title"], cd["Phone"], cd["Email"],
                                cd["Street"], cd["City"], cd["State"] || 'CA', cd["Zip"], new Date()]);
                }
            }

            // === CLAIMANT ===
            if (cd["Claimant Name"]) {
                const s = ss.getSheetByName("Claimants");
                if (s) {
                    const id = generateId_(s);
                    ids.claimantId = id;
                    s.appendRow([id, cd["Claimant Name"], cd["Job Title"], cd["Phone"], cd["Email"],
                                cd["SSN"], cd["Date of Birth"], cd["Street"], cd["City"], cd["State"] || 'CA',
                                cd["Zip"], new Date()]);
                }
            }

            // === CASE ===
            const s = ss.getSheetByName("Cases");
            if (!s) throw new Error("Cases sheet not found");
            const id = generateId_(s);
            recordId = id;

            s.appendRow([
                id,
                cd["Claim Number"], ids.clientId || '', ids.attorneyId || '',
                ids.employerId || '', ids.claimantId || '', cd["Investigator"],
                cd["Service"], cd["Injury Date"], cd["Assignment Date"] || new Date(),
                        cd["Due Date"], cd["Decision Date"], cd["Status"] || 'Open',
                        cd["Pipeline Stage"], cd["Claimant Statement"], cd["Claimant Working"],
                        cd["Medical Release"], cd["Restrictions"], cd["Injury Description"],
                        cd["Adjuster Instructions"], new Date(), ''
            ]);
        }

        return ContentService.createTextOutput(JSON.stringify({
            status: "success",
            id: recordId
        })).setMimeType(ContentService.MimeType.JSON);

    } catch (err) {
        return ContentService.createTextOutput(JSON.stringify({
            status: "error",
            message: err.toString()
        })).setMimeType(ContentService.MimeType.JSON);
    }
}
