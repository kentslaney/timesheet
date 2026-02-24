/**
 * Google Sheets API integration
 * Handles OAuth via Google Identity Services and spreadsheet creation via Sheets API v4
 */

const SCOPES = 'https://www.googleapis.com/auth/spreadsheets https://www.googleapis.com/auth/drive.file https://www.googleapis.com/auth/script.projects';
const SHEETS_API = 'https://sheets.googleapis.com/v4/spreadsheets';
const APPS_SCRIPT_API = 'https://script.googleapis.com/v1/projects';

let tokenClient = null;
let accessToken = null;

/**
 * Initialize the Google Identity Services token client
 */
function initGoogleAuth(clientId, onSignIn, onSignOut) {
    tokenClient = google.accounts.oauth2.initTokenClient({
        client_id: clientId,
        scope: SCOPES,
        callback: (tokenResponse) => {
            if (tokenResponse.error) {
                console.error('Auth error:', tokenResponse.error);
                onSignOut();
                return;
            }
            accessToken = tokenResponse.access_token;
            onSignIn(tokenResponse);
        },
    });
}

/**
 * Request sign-in (triggers OAuth consent)
 */
function requestSignIn() {
    if (!tokenClient) throw new Error('Google auth not initialized');
    tokenClient.requestAccessToken({ prompt: 'consent' });
}

/**
 * Sign out by revoking the token
 */
function signOut(callback) {
    if (accessToken) {
        google.accounts.oauth2.revoke(accessToken, () => {
            accessToken = null;
            if (callback) callback();
        });
    }
}

/**
 * Check if currently authenticated
 */
function isAuthenticated() {
    return accessToken !== null;
}

/**
 * Make an authenticated API request
 */
async function apiRequest(url, options = {}) {
    if (!accessToken) throw new Error('Not authenticated');
    const response = await fetch(url, {
        ...options,
        headers: {
            'Authorization': `Bearer ${accessToken}`,
            'Content-Type': 'application/json',
            ...(options.headers || {}),
        },
    });
    if (!response.ok) {
        const error = await response.json();
        throw new Error(error.error?.message || `API error: ${response.status}`);
    }
    return response.json();
}

/**
 * Convert a format description string into a Sheets API color object
 */
function parseFormatColor(formatStr) {
    const lower = (formatStr || '').toLowerCase();
    const colors = {
        'green': { red: 0.85, green: 0.95, blue: 0.85 },
        'red': { red: 0.95, green: 0.85, blue: 0.85 },
        'blue': { red: 0.85, green: 0.85, blue: 0.95 },
        'yellow': { red: 0.98, green: 0.97, blue: 0.80 },
        'orange': { red: 0.98, green: 0.91, blue: 0.80 },
        'purple': { red: 0.93, green: 0.85, blue: 0.97 },
    };
    for (const [name, color] of Object.entries(colors)) {
        if (lower.includes(name)) return color;
    }
    return { red: 0.85, green: 0.95, blue: 0.85 }; // default green
}

/**
 * Create a bound Apps Script project and upload script code
 */
async function createBoundScript(spreadsheetId, scriptCode, title) {
    // Step 1: Create bound script project
    const project = await apiRequest(APPS_SCRIPT_API, {
        method: 'POST',
        body: JSON.stringify({
            title: title + ' Script',
            parentId: spreadsheetId,
        }),
    });

    // Step 2: Upload the script content
    await apiRequest(`${APPS_SCRIPT_API}/${project.scriptId}/content`, {
        method: 'PUT',
        body: JSON.stringify({
            files: [
                {
                    name: 'Code',
                    type: 'SERVER_JS',
                    source: scriptCode,
                },
                {
                    name: 'appsscript',
                    type: 'JSON',
                    source: JSON.stringify({
                        timeZone: Intl.DateTimeFormat().resolvedOptions().timeZone,
                        dependencies: {},
                        exceptionLogging: 'STACKDRIVER',
                        runtimeVersion: 'V8',
                    }),
                },
            ],
        }),
    });

    return project.scriptId;
}

/**
 * Create a spreadsheet from parsed sheetsdk spec
 */
async function createSpreadsheet(parsed, title, scriptCode) {
    const sheetProperties = parsed.sheets.map((sheet, i) => ({
        properties: {
            sheetId: i,
            title: sheet.name,
            index: i,
        },
    }));

    // Step 1: Create the spreadsheet with sheets
    const createResult = await apiRequest(SHEETS_API, {
        method: 'POST',
        body: JSON.stringify({
            properties: { title },
            sheets: sheetProperties,
        }),
    });

    const spreadsheetId = createResult.spreadsheetId;

    // Step 2: Set values and formulas via batchUpdate
    const requests = [];
    const valueRanges = [];

    for (let si = 0; si < parsed.sheets.length; si++) {
        const sheet = parsed.sheets[si];
        const sheetId = si;

        // Values
        for (const v of sheet.values) {
            let userEnteredValue;
            if (v.type === 'date') {
                // Google Sheets serial date: days since Dec 30, 1899
                const parts = v.value.split('-');
                const date = new Date(Date.UTC(parseInt(parts[0]), parseInt(parts[1]) - 1, parseInt(parts[2])));
                const epoch = new Date(Date.UTC(1899, 11, 30));
                const serial = Math.floor((date - epoch) / (24 * 60 * 60 * 1000));
                userEnteredValue = { numberValue: serial };
            } else if (v.type === 'number') {
                userEnteredValue = { numberValue: v.value };
            } else {
                userEnteredValue = { stringValue: v.value };
            }

            requests.push({
                updateCells: {
                    rows: [{
                        values: [{
                            userEnteredValue,
                            ...(v.type === 'date' ? {
                                userEnteredFormat: {
                                    numberFormat: { type: 'DATE', pattern: 'yyyy-mm-dd' }
                                }
                            } : {}),
                        }],
                    }],
                    fields: v.type === 'date'
                        ? 'userEnteredValue,userEnteredFormat.numberFormat'
                        : 'userEnteredValue',
                    range: {
                        sheetId,
                        startRowIndex: v.cell.row - 1,
                        endRowIndex: v.cell.row,
                        startColumnIndex: v.cell.col,
                        endColumnIndex: v.cell.col + 1,
                    },
                },
            });
        }

        // Formulas
        for (const f of sheet.formulas) {
            requests.push({
                updateCells: {
                    rows: [{
                        values: [{
                            userEnteredValue: { formulaValue: '=' + f.formula },
                        }],
                    }],
                    fields: 'userEnteredValue',
                    range: {
                        sheetId,
                        startRowIndex: f.cell.row - 1,
                        endRowIndex: f.cell.row,
                        startColumnIndex: f.cell.col,
                        endColumnIndex: f.cell.col + 1,
                    },
                },
            });
        }

        // Notes
        // Group notes by cell, multiple notes = paragraphs (per README)
        const notesByCell = {};
        for (const n of sheet.notes) {
            const key = n.cell.a1;
            if (!notesByCell[key]) {
                notesByCell[key] = { cell: n.cell, texts: [] };
            }
            notesByCell[key].texts.push(n.text);
        }
        for (const { cell, texts } of Object.values(notesByCell)) {
            requests.push({
                updateCells: {
                    rows: [{
                        values: [{
                            note: texts.join('\n'),
                        }],
                    }],
                    fields: 'note',
                    range: {
                        sheetId,
                        startRowIndex: cell.row - 1,
                        endRowIndex: cell.row,
                        startColumnIndex: cell.col,
                        endColumnIndex: cell.col + 1,
                    },
                },
            });
        }

        // Conditional formatting
        for (const ch of sheet.conditionalHighlights) {
            const range = {
                sheetId,
                startColumnIndex: ch.range.col,
                endColumnIndex: (ch.range.endCol ?? ch.range.col) + 1,
            };
            if (ch.range.row) range.startRowIndex = ch.range.row - 1;
            if (ch.range.endRow) range.endRowIndex = ch.range.endRow;
            else if (ch.range.row && !ch.range.endRow && !ch.range.endCol) {
                range.endRowIndex = ch.range.row;
            }

            const bgColor = parseFormatColor(ch.format);

            requests.push({
                addConditionalFormatRule: {
                    rule: {
                        ranges: [range],
                        booleanRule: {
                            condition: {
                                type: 'CUSTOM_FORMULA',
                                values: [{ userEnteredValue: '=' + ch.formula }],
                            },
                            format: {
                                backgroundColor: bgColor,
                            },
                        },
                    },
                    index: 0,
                },
            });
        }
    }

    // Step 3: Execute batch update
    if (requests.length > 0) {
        await apiRequest(`${SHEETS_API}/${spreadsheetId}:batchUpdate`, {
            method: 'POST',
            body: JSON.stringify({ requests }),
        });
    }

    // Step 4: Attach Apps Script if code was provided
    if (scriptCode) {
        try {
            await createBoundScript(spreadsheetId, scriptCode, title);
        } catch (err) {
            showStatus('Apps Script not attached: ' + err.message, 'error');
        }
    }

    return {
        spreadsheetId,
        url: `https://docs.google.com/spreadsheets/d/${spreadsheetId}`,
    };
}
