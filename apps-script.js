/**
 * 2026 Masters Pool — Pick Submission Backend
 *
 * Deploy: Extensions → Apps Script → Deploy → Web App
 *   Execute as: Me
 *   Who has access: Anyone
 *
 * Bound to master sheet: 1gslshCjm_dhkqqYOtxVN_xTXU7k_EreFylWmyAvx2lM
 */

const SHEET_ID = '1gslshCjm_dhkqqYOtxVN_xTXU7k_EreFylWmyAvx2lM';
const ENTRIES_TAB = 'Entries';

function getEntriesSheet() {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  let sheet = ss.getSheetByName(ENTRIES_TAB);
  if (!sheet) {
    sheet = ss.insertSheet(ENTRIES_TAB);
    sheet.appendRow([
      'Entry #', 'Name', 'Email', 'Timestamp',
      'Tier 1 - Pick 1', 'Tier 1 - Pick 2', 'Tier 1 - Pick 3',
      'Tier 2 - Pick 1',
      'Tier 3 - Pick 1', 'Tier 3 - Pick 2'
    ]);
  }
  return sheet;
}

function findExistingEntry(email) {
  const sheet = getEntriesSheet();
  const data = sheet.getDataRange().getValues();
  const emailLower = email.toLowerCase().trim();

  for (let i = 1; i < data.length; i++) {
    if (String(data[i][2]).toLowerCase().trim() === emailLower) {
      return {
        entryNumber: data[i][0],
        name: data[i][1],
        email: data[i][2],
        timestamp: data[i][3],
        picks: {
          tier1: [data[i][4], data[i][5], data[i][6]],
          tier2: [data[i][7]],
          tier3: [data[i][8], data[i][9]]
        }
      };
    }
  }
  return null;
}

/**
 * GET handler — check if email has existing picks
 * Usage: ?action=check&email=user@example.com
 */
function doGet(e) {
  const params = e.parameter;
  const action = params.action || 'check';
  const email = params.email || '';

  if (action === 'check' && email) {
    const existing = findExistingEntry(email);
    if (existing) {
      return jsonResponse({ status: 'exists', entry: existing });
    } else {
      return jsonResponse({ status: 'new' });
    }
  }

  return jsonResponse({ status: 'error', message: 'Missing email parameter' });
}

/**
 * POST handler — submit picks
 * Body JSON: { name, email, tier1: [3 names], tier2: [1 name], tier3: [2 names] }
 */
function doPost(e) {
  try {
    const body = JSON.parse(e.postData.contents);
    const { name, email, tier1, tier2, tier3 } = body;

    // Validate required fields
    if (!name || !email) {
      return jsonResponse({ status: 'error', message: 'Name and email are required' });
    }

    // Validate pick counts
    if (!tier1 || tier1.length !== 3) {
      return jsonResponse({ status: 'error', message: 'Must pick exactly 3 from Tier 1' });
    }
    if (!tier2 || tier2.length !== 1) {
      return jsonResponse({ status: 'error', message: 'Must pick exactly 1 from Tier 2' });
    }
    if (!tier3 || tier3.length !== 2) {
      return jsonResponse({ status: 'error', message: 'Must pick exactly 2 from Tier 3' });
    }

    // Check for duplicate email
    const existing = findExistingEntry(email);
    if (existing) {
      return jsonResponse({
        status: 'duplicate',
        message: 'Picks already submitted for this email',
        entry: existing
      });
    }

    // Write to sheet
    const sheet = getEntriesSheet();
    const lastRow = sheet.getLastRow();
    const entryNumber = lastRow; // Row 1 is header, so entry # = lastRow
    const timestamp = new Date().toISOString();

    sheet.appendRow([
      entryNumber,
      name,
      email,
      timestamp,
      tier1[0], tier1[1], tier1[2],
      tier2[0],
      tier3[0], tier3[1]
    ]);

    return jsonResponse({
      status: 'success',
      message: 'Picks submitted successfully',
      entry: {
        entryNumber,
        name,
        email,
        timestamp,
        picks: { tier1, tier2, tier3 }
      }
    });

  } catch (err) {
    return jsonResponse({ status: 'error', message: err.toString() });
  }
}

function jsonResponse(data) {
  return ContentService
    .createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}
