/**
 * SheetSDK Parser
 * Parses JS files with "use sheetsdk" directive into structured spreadsheet specs.
 */

/**
 * Parse a cell reference like "A1", "B3", "A10:A15", "E2:E"
 * Returns { col, row } or { col, row, endCol, endRow } for ranges
 */
function parseCellRef(ref) {
    const rangeMatch = ref.match(/^([A-Z]+)(\d+)?:([A-Z]+)(\d+)?$/);
    if (rangeMatch) {
        return {
            col: colToIndex(rangeMatch[1]),
            row: rangeMatch[2] ? parseInt(rangeMatch[2]) : null,
            endCol: colToIndex(rangeMatch[3]),
            endRow: rangeMatch[4] ? parseInt(rangeMatch[4]) : null,
            a1: ref,
        };
    }
    const cellMatch = ref.match(/^([A-Z]+)(\d+)$/);
    if (cellMatch) {
        return {
            col: colToIndex(cellMatch[1]),
            row: parseInt(cellMatch[2]),
            a1: ref,
        };
    }
    return null;
}

/**
 * Convert column letter(s) to 0-based index. A=0, B=1, ..., Z=25, AA=26
 */
function colToIndex(col) {
    let idx = 0;
    for (let i = 0; i < col.length; i++) {
        idx = idx * 26 + (col.charCodeAt(i) - 64);
    }
    return idx - 1; // 0-based
}

/**
 * Convert 0-based column index to letter(s)
 */
function indexToCol(idx) {
    let col = '';
    idx += 1; // 1-based
    while (idx > 0) {
        idx--;
        col = String.fromCharCode(65 + (idx % 26)) + col;
        idx = Math.floor(idx / 26);
    }
    return col;
}

/**
 * Validate that a file starts with the required sheetsdk directive
 */
function validateHeader(content) {
    const lines = content.split('\n').map(l => l.trim()).filter(l => l.length > 0);
    let idx = 0;

    // Optional "use strict"
    if (idx < lines.length && (lines[idx] === '"use strict"' || lines[idx] === "'use strict'")) {
        idx++;
    }

    // Required "use sheetsdk"
    if (idx < lines.length && (lines[idx] === '"use sheetsdk"' || lines[idx] === "'use sheetsdk'")) {
        return true;
    }

    return false;
}

/**
 * Parse a value string into a typed value
 */
function parseValue(str) {
    // Date: YYYY-MM-DD
    if (/^\d{4}-\d{2}-\d{2}$/.test(str)) {
        return { type: 'date', value: str };
    }
    // Number
    if (/^-?\d+(\.\d+)?$/.test(str)) {
        return { type: 'number', value: parseFloat(str) };
    }
    // Quoted string — remove surrounding quotes
    if ((str.startsWith('"') && str.endsWith('"')) || (str.startsWith("'") && str.endsWith("'"))) {
        return { type: 'string', value: str.slice(1, -1) };
    }
    // Plain string
    return { type: 'string', value: str };
}

/**
 * Main parser: takes file content string, returns parsed spec
 */
function parseSheetSDK(content) {
    if (!validateHeader(content)) {
        throw new Error('Invalid file: must start with "use sheetsdk" (optionally preceded by "use strict")');
    }

    const lines = content.split('\n');
    const sheets = [];
    let currentSheet = null;
    let currentSection = null;
    let currentVersion = null;
    let currentHighlightFormat = null;
    let inMultiLineComment = false;
    let multiLineBuffer = null;
    let multiLineRef = null;

    for (let i = 0; i < lines.length; i++) {
        const line = lines[i];
        const trimmed = line.trim();

        // Handle multi-line comment continuation
        if (inMultiLineComment) {
            if (trimmed.endsWith('*/')) {
                // End of multi-line comment
                const closingContent = trimmed.slice(0, -2).trim();
                if (closingContent) {
                    multiLineBuffer.push(closingContent);
                }

                // Process the accumulated multi-line content
                if (multiLineRef && currentSection) {
                    processMultiLine(currentSheet, currentSection, multiLineRef, multiLineBuffer);
                }

                inMultiLineComment = false;
                multiLineBuffer = null;
                multiLineRef = null;
                continue;
            }
            if (multiLineBuffer !== null) {
                multiLineBuffer.push(trimmed);
            }
            continue;
        }

        // Sheet block start: /*** sheetname ***/
        const sheetMatch = trimmed.match(/^\/\*\*\*\s+(.+?)\s+\*\*\*\/$/);
        if (sheetMatch) {
            const name = sheetMatch[1];
            if (name === 'EOF') {
                // End of current sheet block
                if (currentSheet) {
                    sheets.push(currentSheet);
                    currentSheet = null;
                    currentSection = null;
                    currentVersion = null;
                }
                continue;
            }
            currentSheet = {
                name,
                isPrimary: false,
                formulas: [],
                values: [],
                notes: [],
                conditionalHighlights: [],
                actions: [],
            };
            currentSection = null;
            currentVersion = null;
            continue;
        }

        if (!currentSheet) continue;

        // Primary/secondary sheet marker: /* primary sheet */ or /* secondary sheet */
        if (/^\/\*\s+(primary|secondary)\s+sheet\s+\*\/$/.test(trimmed)) {
            currentSheet.isPrimary = trimmed.includes('primary');
            continue;
        }

        // Version + section on same line: /** 0.1 **/ /* actions */
        // or version + section + format: /** 0.1 **/ /* conditional highlights */ // "Green background"
        const versionSectionMatch = trimmed.match(
            /^\/\*\*\s+([\d.]+)\s+\*\*\/\s+\/\*\s+(.+?)\s+\*\/(?:\s+\/\/\s+(.+))?$/
        );
        if (versionSectionMatch) {
            currentVersion = versionSectionMatch[1];
            currentSection = versionSectionMatch[2];
            currentHighlightFormat = versionSectionMatch[3]
                ? versionSectionMatch[3].replace(/^"(.*)"$/, '$1')
                : null;
            continue;
        }

        // Section header: /* sectiontype */
        const sectionMatch = trimmed.match(/^\/\*\s+(.+?)\s+\*\/$/);
        if (sectionMatch) {
            currentSection = sectionMatch[1];
            currentVersion = null;
            currentHighlightFormat = null;
            continue;
        }

        // Multi-line comment start (for values or notes): /* A1:E1
        const multiStartMatch = trimmed.match(/^\/\*\s+([A-Z]+\d*:[A-Z]*\d*)$/);
        if (multiStartMatch) {
            inMultiLineComment = true;
            multiLineRef = multiStartMatch[1];
            multiLineBuffer = [];
            continue;
        }

        // Single-line comment entries: // CellRef=content
        const entryMatch = trimmed.match(/^\/\/\s+([A-Z]+\d*(?::[A-Z]+\d*)?)\s*=\s*(.*)$/);
        if (entryMatch && currentSection) {
            const ref = entryMatch[1];
            const content = entryMatch[2];
            const parsed = parseCellRef(ref);

            if (!parsed) continue;

            switch (currentSection) {
                case 'actions':
                    currentSheet.actions.push({ cell: parsed, name: content });
                    break;
                case 'formulas':
                    currentSheet.formulas.push({ cell: parsed, formula: content });
                    break;
                case 'values':
                    currentSheet.values.push({ cell: parsed, ...parseValue(content) });
                    break;
                case 'notes':
                    // Notes can have multiple entries for the same cell (paragraphs)
                    currentSheet.notes.push({ cell: parsed, text: content });
                    break;
                case 'conditional highlights':
                    currentSheet.conditionalHighlights.push({
                        range: parsed,
                        formula: content,
                        format: currentHighlightFormat || 'Green background',
                        version: currentVersion,
                    });
                    break;
            }
            continue;
        }
    }

    // If file ends without /*** EOF ***/, push any remaining sheet
    if (currentSheet) {
        sheets.push(currentSheet);
    }

    return { sheets };
}

/**
 * Process multi-line comment block for values or notes
 */
function processMultiLine(sheet, section, refStr, lines) {
    const parsed = parseCellRef(refStr);
    if (!parsed) return;

    switch (section) {
        case 'values': {
            // Multi-line values: each line is a cell value in the range
            // Range is 1 row or 1 column per README
            const isRow = parsed.row === (parsed.endRow || parsed.row);
            lines.forEach((line, idx) => {
                const cellRef = isRow
                    ? { col: parsed.col + idx, row: parsed.row, a1: indexToCol(parsed.col + idx) + parsed.row }
                    : { col: parsed.col, row: parsed.row + idx, a1: indexToCol(parsed.col) + (parsed.row + idx) };
                sheet.values.push({ cell: cellRef, ...parseValue(line) });
            });
            break;
        }
        case 'notes': {
            // Multi-line notes: each line corresponds to a cell in the range
            const isRow = parsed.row === (parsed.endRow || parsed.row);
            lines.forEach((line, idx) => {
                const cellRef = isRow
                    ? { col: parsed.col + idx, row: parsed.row, a1: indexToCol(parsed.col + idx) + parsed.row }
                    : { col: parsed.col, row: parsed.row + idx, a1: indexToCol(parsed.col) + (parsed.row + idx) };
                sheet.notes.push({ cell: cellRef, text: line });
            });
            break;
        }
    }
}

/**
 * Extract the Apps Script code from a sheetsdk JS file.
 * Strips the header directives and all spec blocks (delimited by triple-star comments),
 * returning only the executable JavaScript code.
 */
function extractScriptCode(content) {
    const lines = content.split('\n');
    const outputLines = [];
    let inSpecBlock = false;
    let inMultiLineComment = false;
    let headerDone = false;

    for (let i = 0; i < lines.length; i++) {
        const trimmed = lines[i].trim();

        // Skip header directives
        if (!headerDone) {
            if (trimmed === '' || trimmed === '"use strict"' || trimmed === "'use strict'" ||
                trimmed === '"use sheetsdk"' || trimmed === "'use sheetsdk'") {
                continue;
            }
            headerDone = true;
        }

        // Track multi-line comments inside spec blocks
        if (inSpecBlock && inMultiLineComment) {
            if (trimmed.endsWith('*/')) {
                inMultiLineComment = false;
            }
            continue;
        }

        // Sheet block start/end: /*** name ***/
        const sheetMatch = trimmed.match(/^\/\*\*\*\s+(.+?)\s+\*\*\*\/$/);
        if (sheetMatch) {
            if (sheetMatch[1] === 'EOF') {
                inSpecBlock = false;
            } else {
                inSpecBlock = true;
            }
            continue;
        }

        if (inSpecBlock) {
            // Track multi-line comment starts within spec blocks
            if (trimmed.startsWith('/*') && !trimmed.endsWith('*/')) {
                inMultiLineComment = true;
            }
            continue;
        }

        outputLines.push(lines[i]);
    }

    // Trim leading/trailing blank lines
    let start = 0;
    while (start < outputLines.length && outputLines[start].trim() === '') start++;
    let end = outputLines.length - 1;
    while (end >= 0 && outputLines[end].trim() === '') end--;

    return outputLines.slice(start, end + 1).join('\n');
}

// Export for Node.js testing or keep in global scope for browser
if (typeof module !== 'undefined' && module.exports) {
    module.exports = { parseSheetSDK, validateHeader, parseCellRef, colToIndex, indexToCol, parseValue, extractScriptCode };
}
