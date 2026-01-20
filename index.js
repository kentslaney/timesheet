"use strict"
"use drivesdk"

/*** timer ***/
/* primary sheet */

/** 0.1 **/ /* buttons */
// A6=mark
// A7=newNotes
// A8=dupNotes

/* formulas */
// B3=LAMBDA(t,IF(AND(TODAY()=B1,$A$1>B2),QUOTIENT(MOD(t,24*60),60)*100+MOD(t, 60),""))(QUOTIENT($A$1, 100)*60+MOD($A$1,100)-(SUM(BYROW({B4:B,ARRAYFORMULA(ROW(B4:B))},LAMBDA(row,(QUOTIENT(INDEX(row,1), 100)*60+MOD(INDEX(row,1),100))*IF(MOD(INDEX(row,2),2),1,-1))))-IF(MOD(COUNTA(B4:B),2),,QUOTIENT($A$5, 100)*60+MOD($A$5,100))))
// B2=LAMBDA(t,QUOTIENT(t,60)*100+MOD(t, 60))(SUM(BYROW({B4:B,ARRAYFORMULA(ROW(B4:B))},LAMBDA(row,(QUOTIENT(INDEX(row,1), 100)*60+MOD(INDEX(row,1),100))*IF(MOD(INDEX(row,2),2),1,-1))))+IF(MOD(COUNTA(B4:B),2),IF(TODAY()=B1,QUOTIENT($A$5, 100)*60+MOD($A$5,100),24*60)))
// A5=MINUTE(NOW()) + (HOUR(NOW()) - HOUR(TODAY())) * 100
// A9=IF(AND(A7="toggle", A8="update"),,) + AND(TODAY()=INDIRECT("B1"),MOD(COUNTA(INDIRECT("B4:B")),2))
// A10=ARRAYFORMULA({"1kPUeEVGR-eUKYQJGZ82z9zolk8W11mIStZcMzWJvtlI"; "toggle"; "update"; "mark"; "notes"; "duplicate"})

/* values */
// A1=800
// A2="total"
// A3="earliest end"
// A4="ToD"
// B1=1970-01-01

/* notes */
// A1=goal time per day (for earliest end)
// A9=1 if the timer is running, 0 otherwise. B1 is also conditionally formatted to be green when this is 1. Running across midnight will delay day column creation until user action.
// A9=(first expression is to add a dependency to the placeholder cells)
/* A10:A15
Notes doc template id (text after /d/ in the URL until the next slash character)
edit to toggle timer with current time and note given
edit to keep timer stopped/running, but mark the time with the given note
toggle timer without an explanation (any edit triggers action then resets cell)
create a duplicate of the template sheet in A10 (any edit triggers action then resets cell)
create a duplicate of the previous day's notes in C3 (any edit triggers action then resets cell)
*/
/*** EOF ***/


function occupied(sheet, ...range) {
  return sheet.getRange(...range).getValues().filter(x => x != "").length;
}

function lastOccupied(sheet, col) {
  return sheet.getRange(1, col, sheet.getLastRow(), 1).getValues().map(x => x != "").lastIndexOf(true) + 1;
}

function daysDiff(target, until) {
  const [start, end] = until === undefined ? [target, new Date()] : [target, until];
  return Math.trunc((end.getTime() - start.getTime()) / (1000 * 60 * 60 * 24));
}

function coerce(id) {
  if (id === undefined) {
    return SpreadsheetApp.getActiveSheet();
  } else if (id instanceof Number) {
    let ss = SpreadsheetApp.getActiveSpreadsheet();
    let sheet = ss.getSheets().find(s => s.getSheetId() == id);
    if (sheet === undefined)
      throw new Error("no sheet with that id");
    return sheet;
  } else if (id.hasOwnProperty("getSheet")) {
    return id.getSheet();
  } else {
    return id;
  }
}

function dateFormat(date) {
  return date.getHours() * 100 + date.getMinutes();
}

function assureDate(sheet, date) {
  sheet = sheet || SpreadsheetApp.getActiveSheet();
  const latest = sheet.getRange("B1").getValue();
  if (!(latest instanceof Date))
    throw new Error("B1 is not a date");
  const when = date === undefined ? new Date() : date;
  if (daysDiff(latest, when) > 0) {
    sheet.insertColumnAfter(1);
    sheet.getRange("C1:C3").copyTo(sheet.getRange("B1:B3"));
    sheet.getRange("C4:C").copyTo(sheet.getRange("B4:B"), {formatOnly: true});
    const todayStart = new Date(when.getFullYear(), when.getMonth(), when.getDate());
    sheet.getRange("B1").setValue(todayStart);
    if(daysDiff(latest, when) === 1 && occupied(sheet, "C4:C") % 2)
      sheet.getRange("B4").setValue(0);
    if (!(sheet.getRange("D1").getValue() instanceof Date) &&
        lastOccupied(sheet, 4) == 3 && occupied(sheet, "D:D") == 1) {
      sheet.getRange("D3").copyTo(sheet.getRange("C3"));
      sheet.deleteColumn(4);
    }
    fixFormat(sheet);
  }
}

function mark(id) {
  let sheet = coerce(id);
  const now = new Date();
  assureDate(sheet, now);
  const row = 4 + occupied(sheet, "B4:B");
  sheet.getRange(row, 2, 1, 1).setValue(dateFormat(now));
  return row;
}

function cellDefault(sheet, id, message) {
  if (message !== undefined)
    return message;
  if (id instanceof sheet.getRange(1, 1).constructor)
    return id.getValue();
  throw new Error("invalid message")
}

function assureNotes(sheet) {
  if (sheet.getRange("C1").getValue() instanceof Date)
    sheet.insertColumnBefore(3);
}

function note(id, message) {
  let sheet = coerce(id);
  let valued = cellDefault(sheet, id, message);
  if (message === undefined && valued === "") return;
  const row = mark(sheet);
  assureNotes(sheet);
  sheet.getRange(row, 3, 1, 1).setValue(valued);
  if (message === undefined)
    id.setValue("");
  return row;
}

function update(id, message) {
  let sheet = coerce(id);
  let valued = cellDefault(sheet, id, message);
  if (message === undefined && valued === "") return;
  let row = note(sheet, valued);
  sheet.getRange(row, 2, 1, 1).copyTo(sheet.getRange(row + 1, 2, 1, 1));
  sheet.getRange(row, 3, 2, 1).mergeVertically();
  if (message === undefined) id.setValue("");
}


function onEdit(e) {
  const range = e.range;

  if (range.getSheet().getSheetId() == 0) {
    switch (range.getA1Notation()) {
      case "A11": note(range); break;
      case "A12": update(range); break;
      case "A13": mark(); range.setValue(""); break;
      case "A14": newNotes(); range.setValue(""); break;
      case "A15": dupNotes(); range.setValue(""); break;
    }
  }
}

function datedCWD() {
  let parents = DriveApp.getFileById(SpreadsheetApp.getActive().getId()).getParents();
  let folder = parents.next();
  if (parents.hasNext())
    throw new Error("unable to create copy with multiple parents");
  let name = new Date(new Date() + " UTC").toISOString().split('T')[0]; // lol
  return [name, folder]
}

function deduplicate(name, folder) {
  let dups = folder.getFilesByName(name);
  if (dups.hasNext()) {
    let ui = SpreadsheetApp.getUi();
    let response = ui.alert('Create duplicated name',
      "There is already a document named with today's date. Would you like to create another?",
      ui.ButtonSet.YES_NO);
    if (response == ui.Button.NO) return dups.next().getUrl();
  }
}

function notesDoc(template) {
  template = DriveApp.getFileById(IdURL(template));
  let file = datedCWD()
  let existing = deduplicate(...file)
  if (existing === undefined) return template.makeCopy(...file).getUrl();
  else return existing
}

// https://developers.google.com/docs/api/concepts/document#document-id
const IdRegExp = new RegExp("/document/d/([a-zA-Z0-9-_]+)", "g")

function IdURL(url) {
  if (/^[a-zA-Z0-9-_]+$/.test(url)) return url
  let {done, value} = url.matchAll(IdRegExp).next()
  return done ? null : value[1]
}

function dayNotes(a1str) {
  let sheet = SpreadsheetApp.getActiveSheet();
  let url = notesDoc(sheet.getRange(...(typeof a1str === "string" ? [a1str] : a1str)).getDisplayValue());
  assureDate(sheet);
  assureNotes(sheet);
  sheet.getRange("C3").setValue(url);
}

function newNotes() { dayNotes("A10") }
function dupNotes() {
  let sheet = SpreadsheetApp.getActiveSheet();
  let values = sheet.getRange("C3:3").getDisplayValues()
  let idx = values[0].map(x => Boolean(x)).indexOf(true)
  dayNotes([3, 3 + idx])
}

function fixFormat(id) {
  let sheet = coerce(id);
  var rules = sheet.getConditionalFormatRules();
  const relevant = [2, 2]
  let rules_updated = false;

  rules = rules.map(x => {
    let range_updated = false;
    let fixed = x.getRanges().reduce((i, j) => {
      const bound = [[j.getRow(), j.getColumn()],
        [j.getLastRow(), j.getLastColumn()]];
      let replacement = []
      if (bound[0][0] <= relevant[0] && bound[1][0] >= relevant[0] &&
          bound[1][1] > relevant[1]) {
        if (bound[0][0] < relevant[0])
          replacement.push([bound[0][0], bound[0][1],
            relevant[0] - bound[0][0] - 1, bound[1][1] - bound[0][1]]);
        if (bound[0][1] <= relevant[1])
          replacement.push([relevant[0], bound[0][1],
            1, relevant[1] - bound[0][1] + 1]);
        if (bound[1][0] > relevant[0])
          replacement.push([relevant[0] + 1, bound[0][1],
            bound[0][0] - relevant[0] - 1, bound[1][1] - bound[0][1]]);
        rules_updated = range_updated = true;
        return i.concat(replacement.map(x => sheet.getRange(...x)));
      }
      return i.concat([j])
    }, [])
    return range_updated ? x.copy().setRanges(fixed) : x;
  })

  if (rules_updated) sheet.setConditionalFormatRules(rules);
}

function addMinutes(date, minutes) {
    return new Date(date.getTime() + minutes*60000);
}

function local_tz(init) {
  let obj = new Date(init)
  return addMinutes(obj, obj.getTimezoneOffset())
}

const MS24H = 24 * 60 * 60 * 1000

function listRevisions(fileId) {
  // DETOUR (kentslaney) - Could be a generator
  let timestamps = [];
  let revisions;
  let pageToken = null;
  do {
    try {
      revisions = Drive.Revisions.list(
          fileId,
          {'fields': 'revisions(modifiedTime,size,id),nextPageToken'});
      if (!revisions.revisions || revisions.revisions.length === 0) {
        console.log('All revisions found.');
        return;
      }
      for (let i = 0; i < revisions.revisions.length; i++) {
        const revision = revisions.revisions[i];
        const date = new Date(revision.modifiedTime);
        timestamps.push(date)
      }
      pageToken = revisions.nextPageToken;
    } catch (err) {
      // TODO (kentslaney) - Handle exception
      console.log('Failed with error %s', err.message);
    }
  } while (pageToken);
  return timestamps
}

function filterRevisionDay(fileId, day) {
  const revisions = listRevisions(fileId)
  const isDay = revisions.map(x => x.getTime() - day).filter(x => x >= 0 && x <= MS24H)
  return revisions.filter((_, i) => isDay[i])
}

function editTimeBounds(fileId, day) {
  const revisions = filterRevisionDay(fileId, day)
  if (revisions.length === 0) return [-2400, -2400]
  const start = dateFormat(revisions[0])
  const end = dateFormat(revisions[revisions.length - 1])
  return [-end, -start]
}

function estimateTimes(sheet, c1) {
  return editTimeBounds(IdURL(sheet.getRange(3, c1).getValue()), sheet.getRange(1, c1).getValue())
}

function offerTimes(c1) {
  let sheet = SpreadsheetApp.getActiveSheet();
  const [[a], [_], [c], [d], [e]] = sheet.getRange(1, c1, 5, 1).getValues()
  if (!(a instanceof Date) || IdURL(c) === null || d || e) return;
  const [lo, hi] = estimateTimes(sheet, c1)
  sheet.getRange(4, c1).setValue(lo)
  sheet.getRange(5, c1).setValue(hi)
}

/*** parsed ***/
/* formulas */
// A2=ARRAYFORMULA(TRANSPOSE(FILTER(INDIRECT("timer!B1:1"),NOT(ISBLANK(INDIRECT("timer!B2:2"))))))
// B2=ARRAYFORMULA(TRANSPOSE(LAMBDA(x,QUOTIENT(x, 100)*60+MOD(x, 100))(FILTER(INDIRECT("timer!B2:2"),NOT(ISBLANK(INDIRECT("timer!B2:2")))))))
// C2=ARRAYFORMULA(IF(ISBLANK(B2:B),,B2:B/60))
// D2=BYROW(A2:A,LAMBDA(i, IF(ISBLANK(i),,LAMBDA(j, SUM(FILTER(INDEX(j,0,2),i-INDEX(j,0,1)<7))/60)(INDIRECT("R" & ROW(i) & "C1:C2", false)))))
// E2=BYROW(A2:A,LAMBDA(i, IF(ISBLANK(i),,LAMBDA(j, SUM(FILTER(INDEX(j,0,2),INDEX(j,0,1)>=DATE(YEAR(i),MONTH(i),DAY(i)-MOD(WEEKDAY(i-2),7))))/60)(INDIRECT("R" & ROW(i) & "C1:C2", false)))))
// F1=SUM(B2:B)/(MAX(A2:A)-MIN(A2:A)+1)/60*7/5

/** 0.1 **/ /* conditional highlights */ // "Green background"
// E2:E=IF(ISBLANK(A2),false,IF(ROW()=2,true,OR(ARRAYFORMULA(WEEKDAY(SEQUENCE(DATEDIF(A2,A1,"D"),1,A2+1))=2))))

/* values */
/* A1:E1
date
minutes
hours
hrs last 7 days
week's total hrs
*/

/* notes */
// F1=average hours per work day
/*** EOF ***/
