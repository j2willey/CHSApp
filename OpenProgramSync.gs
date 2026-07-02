/**
 * CHS Open Program → Firestore Sync
 *
 * Whenever a cell is edited in this sheet, the corresponding
 * session/day document in Firestore is updated automatically.
 *
 * SETUP (one-time):
 *  1. Extensions → Apps Script
 *  2. Paste this entire file into Code.gs
 *  3. Click the gear icon (Project Settings) → Script Properties
 *  4. Add a property named FIREBASE_SERVICE_ACCOUNT
 *     Value = the full JSON text of your Firebase service account key
 *     (Firebase Console → Project Settings → Service Accounts → Generate new private key)
 *  5. Back in the editor: Triggers (clock icon) → Add Trigger
 *     Function: onSheetEdit | Event source: Spreadsheet | Event type: On edit
 *  6. Also add a second trigger:
 *     Function: onSheetChange | Event source: Spreadsheet | Event type: On change
 *     (catches row insertions / deletions)
 *  7. Click Save, authorize when prompted.
 *
 * To do a full manual sync at any time, run syncAll() from the editor.
 */

const PROJECT_ID = 'camphisierraapp';
const FIRESTORE_BASE =
  `https://firestore.googleapis.com/v1/projects/${PROJECT_ID}/databases/(default)/documents`;

// ─── Trigger handlers ────────────────────────────────────────────────────────

/** Fires on individual cell edits — only syncs the affected session/day. */
function onSheetEdit(e) {
  try {
    const sheet = e.range.getSheet();
    const allData = sheet.getDataRange().getValues();
    const firstRow = e.range.getRow();
    const lastRow  = e.range.getLastRow();

    const affectedKeys = new Set();
    for (let r = firstRow; r <= lastRow; r++) {
      if (r === 1) continue; // skip header
      const row = allData[r - 1];
      const session = String(row[0]).trim();
      const day     = String(row[1]).trim();
      if (session && day) affectedKeys.add(`${session}|${day}`);
    }

    if (affectedKeys.size === 0) return;

    const grouped = groupBySessionDay(allData.slice(1));
    const token   = getServiceAccountToken();

    for (const key of affectedKeys) {
      if (grouped[key]) {
        const { session, day, steps } = grouped[key];
        updateFirestore(token, session, day, steps);
      }
    }
  } catch (err) {
    Logger.log('onSheetEdit error: ' + err.message);
  }
}

/** Fires on structural changes (row insert/delete) — does a full sync. */
function onSheetChange(e) {
  try {
    syncAll();
  } catch (err) {
    Logger.log('onSheetChange error: ' + err.message);
  }
}

// ─── Core sync ───────────────────────────────────────────────────────────────

/** Syncs every session/day in the sheet to Firestore. Run this manually to push all changes. */
function syncAll() {
  const sheet   = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const allData = sheet.getDataRange().getValues();
  const grouped = groupBySessionDay(allData.slice(1));
  const token   = getServiceAccountToken();

  const keys = Object.keys(grouped);
  Logger.log(`Syncing ${keys.length} session/day combos to Firestore…`);

  for (const key of keys) {
    const { session, day, steps } = grouped[key];
    updateFirestore(token, session, day, steps);
  }

  Logger.log('Sync complete.');
}

// ─── Helpers ─────────────────────────────────────────────────────────────────

/**
 * Groups sheet rows (already without header) into { "Session|Day": { session, day, steps[] } }.
 * Rows are assumed to already be in order within each session/day.
 */
function groupBySessionDay(rows) {
  const grouped = {};

  for (const row of rows) {
    const session     = String(row[0]).trim();
    const day         = String(row[1]).trim();
    const area        = String(row[3]).trim();
    const activity    = String(row[4]).trim();
    const description = String(row[5]).trim();

    if (!session || !day || !activity) continue;

    const key = `${session}|${day}`;
    if (!grouped[key]) grouped[key] = { session, day, steps: [] };

    // Reconstruct the step string in the same format the app expects:
    // "Area: Activity — Description"  or  "Activity" (for arrival/closing rows)
    let step = '';
    if (area) {
      step = `${area}: ${activity}`;
      if (description) step += ` — ${description}`; // em dash
    } else {
      step = activity;
      if (description) step += ` — ${description}`;
    }

    grouped[key].steps.push(step);
  }

  return grouped;
}

/**
 * PATCHes only the `steps` field of sessions/{session}/skills/{day} in Firestore.
 * All other fields (title, description, icon, location) are left untouched.
 */
function updateFirestore(token, session, day, steps) {
  const url =
    `${FIRESTORE_BASE}/sessions/${encodeURIComponent(session)}/skills/${encodeURIComponent(day)}` +
    `?updateMask.fieldPaths=steps`;

  const body = {
    fields: {
      steps: {
        arrayValue: {
          values: steps.map(s => ({ stringValue: s }))
        }
      }
    }
  };

  const resp = UrlFetchApp.fetch(url, {
    method: 'PATCH',
    contentType: 'application/json',
    headers: { Authorization: `Bearer ${token}` },
    payload: JSON.stringify(body),
    muteHttpExceptions: true
  });

  const code = resp.getResponseCode();
  if (code !== 200) {
    Logger.log(`Firestore PATCH ${session}/${day} → ${code}: ${resp.getContentText()}`);
  } else {
    Logger.log(`Updated ${session}/${day} (${steps.length} steps)`);
  }
}

/**
 * Obtains a short-lived Google OAuth2 access token using the Firebase service account
 * stored in Script Properties as FIREBASE_SERVICE_ACCOUNT (JSON string).
 */
function getServiceAccountToken() {
  const props = PropertiesService.getScriptProperties();
  const raw   = props.getProperty('FIREBASE_SERVICE_ACCOUNT');
  if (!raw) throw new Error('Script property FIREBASE_SERVICE_ACCOUNT is not set.');

  const key = JSON.parse(raw);
  const now = Math.floor(Date.now() / 1000);

  const header = { alg: 'RS256', typ: 'JWT' };
  const claim  = {
    iss:   key.client_email,
    scope: 'https://www.googleapis.com/auth/datastore',
    aud:   'https://oauth2.googleapis.com/token',
    iat:   now,
    exp:   now + 3600
  };

  const toSign =
    Utilities.base64EncodeWebSafe(JSON.stringify(header)) + '.' +
    Utilities.base64EncodeWebSafe(JSON.stringify(claim));

  const signature = Utilities.computeRsaSha256Signature(toSign, key.private_key);
  const jwt = toSign + '.' + Utilities.base64EncodeWebSafe(signature);

  const resp = UrlFetchApp.fetch('https://oauth2.googleapis.com/token', {
    method: 'POST',
    contentType: 'application/x-www-form-urlencoded',
    payload: `grant_type=urn%3Aietf%3Aparams%3Aoauth%3Agrant-type%3Ajwt-bearer&assertion=${jwt}`,
    muteHttpExceptions: true
  });

  const result = JSON.parse(resp.getContentText());
  if (!result.access_token) throw new Error('Failed to get token: ' + resp.getContentText());
  return result.access_token;
}
