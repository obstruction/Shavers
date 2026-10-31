/**
 * Dinners page helper — lives in the "Meal Planner" spreadsheet
 * (Extensions → Apps Script). Lets the Dinners page on GitHub Pages
 * read the week live and change it:
 *   - pick a meal for a day (including from "Try something new")
 *   - change or clear a day's meal
 *   - swap two days (meal and notes move together)
 *
 * It only writes the Meal (column B) and Notes (column E) cells of
 * "Meals schedule". Everything else stays formula-driven.
 *
 * SETUP
 * 1. Change FAMILY_CODE below to something the kids can type.
 * 2. Deploy → New deployment → type "Web app".
 *    Execute as: Me. Who has access: Anyone.
 * 3. Copy the web app URL (ends in /exec) into SCRIPT_URL in the
 *    Dinners page's index.html.
 * After editing this file later: Deploy → Manage deployments → edit
 * (pencil) → Version: New version → Deploy. The URL stays the same.
 */

const FAMILY_CODE = 'change-me';

const SCHEDULE = 'Meals schedule';   // A Date, B Meal, E Notes
const PHONE = 'This Week (phone)';   // what the page shows
const ALL = 'All meals';             // A Name, B Main/Side/Dessert

function doGet() {
  return json_(readWeek_());
}

function doPost(e) {
  let body = {};
  try { body = JSON.parse(e.postData.contents); } catch (err) {}

  if (FAMILY_CODE && String(body.code || '').trim().toLowerCase() !== FAMILY_CODE.toLowerCase()) {
    return json_({ error: 'code' });
  }

  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const sh = SpreadsheetApp.getActive().getSheetByName(SCHEDULE);

    if (body.action === 'set') {
      const r = findRow_(sh, body.date);
      if (!r) return json_({ error: 'That day isn’t in the schedule.' });
      sh.getRange(r, 2).setValue(String(body.meal || ''));
    } else if (body.action === 'swap') {
      const a = findRow_(sh, body.a);
      const b = findRow_(sh, body.b);
      if (!a || !b) return json_({ error: 'That day isn’t in the schedule.' });
      [2, 5].forEach(col => {
        const ca = sh.getRange(a, col), cb = sh.getRange(b, col);
        const va = ca.getValue(), vb = cb.getValue();
        ca.setValue(vb);
        cb.setValue(va);
      });
    } else {
      return json_({ error: 'Unknown action.' });
    }

    SpreadsheetApp.flush();
  } finally {
    lock.releaseLock();
  }
  return json_(readWeek_());
}

function readWeek_() {
  const ss = SpreadsheetApp.getActive();
  const tz = ss.getSpreadsheetTimeZone();
  const ph = ss.getSheetByName(PHONE);
  const n = ph.getLastRow();
  const rng = ph.getRange(3, 1, Math.max(n - 2, 1), 6);
  const vals = rng.getValues();
  const disp = rng.getDisplayValues();

  const days = [];
  const recs = [];
  let inRecs = false;
  for (let i = 0; i < vals.length; i++) {
    const a = disp[i][0].trim();
    if (/^recommender:?$/i.test(a)) { inRecs = true; continue; }
    if (inRecs) {
      if (a && recs.indexOf(a) === -1) recs.push(a);
    } else if (vals[i][0] instanceof Date) {
      days.push({
        iso: Utilities.formatDate(vals[i][0], tz, 'yyyy-MM-dd'),
        label: a,
        meal: disp[i][1], notes: disp[i][2], prep: disp[i][3],
        url: disp[i][4], image: disp[i][5]
      });
    }
  }

  const all = ss.getSheetByName(ALL);
  const meals = all.getRange(3, 1, Math.max(all.getLastRow() - 2, 1), 2).getDisplayValues()
    .filter(r => r[0].trim() !== '')
    .map(r => ({ name: r[0].trim(), main: /main/i.test(r[1]) }));

  return { days: days, recs: recs, meals: meals };
}

function findRow_(sh, iso) {
  if (!iso) return 0;
  const tz = sh.getParent().getSpreadsheetTimeZone();
  const dates = sh.getRange(3, 1, sh.getLastRow() - 2, 1).getValues();
  for (let i = 0; i < dates.length; i++) {
    const d = dates[i][0];
    if (d instanceof Date && Utilities.formatDate(d, tz, 'yyyy-MM-dd') === iso) return i + 3;
  }
  return 0;
}

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
