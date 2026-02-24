/**
 * Parser test — run with: node test-parser.js
 */
const { parseSheetSDK, validateHeader, parseCellRef, colToIndex } = require('./parser.js');
const fs = require('fs');
const path = require('path');

let passed = 0;
let failed = 0;

function assert(condition, msg) {
    if (condition) {
        passed++;
    } else {
        failed++;
        console.error(`  FAIL: ${msg}`);
    }
}

// Test 1: validateHeader
console.log('Test: validateHeader');
assert(validateHeader('"use sheetsdk"\n'), 'bare use sheetsdk');
assert(validateHeader('"use strict"\n"use sheetsdk"\n'), 'use strict + use sheetsdk');
assert(!validateHeader('// hello\n"use sheetsdk"\n'), 'comment before sheetsdk rejected');
assert(!validateHeader('"use strict"\n'), 'use strict alone rejected');
assert(!validateHeader(''), 'empty file rejected');

// Test 2: colToIndex
console.log('Test: colToIndex');
assert(colToIndex('A') === 0, 'A=0');
assert(colToIndex('B') === 1, 'B=1');
assert(colToIndex('Z') === 25, 'Z=25');
assert(colToIndex('AA') === 26, 'AA=26');

// Test 3: parseCellRef
console.log('Test: parseCellRef');
const a1 = parseCellRef('A1');
assert(a1 && a1.col === 0 && a1.row === 1, 'A1 parsed');
const e2e = parseCellRef('E2:E');
assert(e2e && e2e.col === 4 && e2e.row === 2 && e2e.endCol === 4 && e2e.endRow === null, 'E2:E parsed');

// Test 4: Parse timesheet.js
console.log('Test: parse timesheet.js');
const content = fs.readFileSync(path.join(__dirname, 'examples', 'timesheet.js'), 'utf-8');
const result = parseSheetSDK(content);

assert(result.sheets.length === 2, `found 2 sheets (got ${result.sheets.length})`);

const timer = result.sheets[0];
assert(timer.name === 'timer', `first sheet named "timer" (got "${timer.name}")`);
assert(timer.isPrimary === true, 'timer is primary');

// Actions
assert(timer.actions.length === 3, `timer has 3 actions (got ${timer.actions.length})`);
assert(timer.actions[0].name === 'mark', `first action is mark`);

// Formulas
assert(timer.formulas.length >= 5, `timer has >=5 formulas (got ${timer.formulas.length})`);
assert(timer.formulas[0].cell.a1 === 'B3', `first formula is B3`);
assert(timer.formulas[0].formula.startsWith('LAMBDA'), `B3 formula starts with LAMBDA`);

// Values
assert(timer.values.length >= 5, `timer has >=5 values (got ${timer.values.length})`);
const a1val = timer.values.find(v => v.cell.a1 === 'A1');
assert(a1val && a1val.value === 800, `A1=800`);
const b1val = timer.values.find(v => v.cell.a1 === 'B1');
assert(b1val && b1val.type === 'date', `B1 is a date`);

// Notes
assert(timer.notes.length >= 2, `timer has >=2 notes (got ${timer.notes.length})`);

// Second sheet: parsed
const parsed = result.sheets[1];
assert(parsed.name === 'parsed', `second sheet named "parsed" (got "${parsed.name}")`);
assert(parsed.formulas.length >= 5, `parsed has >=5 formulas (got ${parsed.formulas.length})`);
assert(parsed.conditionalHighlights.length === 1, `parsed has 1 conditional highlight (got ${parsed.conditionalHighlights.length})`);
assert(parsed.conditionalHighlights[0].format === 'Green background', `highlight format is "Green background"`);

// Multi-line values
const dateVal = parsed.values.find(v => v.cell.a1 === 'A1');
assert(dateVal && dateVal.value === 'date', `parsed A1 = "date" (got "${dateVal?.value}")`);
const minutesVal = parsed.values.find(v => v.cell.a1 === 'B1');
assert(minutesVal && minutesVal.value === 'minutes', `parsed B1 = "minutes" (got "${minutesVal?.value}")`);
const weeksVal = parsed.values.find(v => v.cell.a1 === 'E1');
assert(weeksVal && weeksVal.value === "week's total hrs", `parsed E1 = "week's total hrs" (got "${weeksVal?.value}")`);

// Summary
console.log(`\n${passed} passed, ${failed} failed`);
process.exit(failed > 0 ? 1 : 0);
