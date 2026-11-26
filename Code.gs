/**
 * Job Search Tracker: Gmail auto-logger + quick-add row.
 *
 * Paste into Extensions > Apps Script in the tracker, save, reload the sheet.
 * Then run "Job tracker > Turn on daily Gmail check" once and approve access.
 *
 * What it does (daily ~7:00, or on demand from the menu):
 *  - Reads recent Gmail about applications (EN + PL).
 *  - Rejection matched to exactly one open row: sets Status to Rejected,
 *    fills "Last company reply", adds a one-line note in Next Action.
 *  - Interview / next-step email: fills "Last company reply", sets Follow-up date
 *    to today so the row turns orange. Status is left for you to change.
 *  - Every email it finds, matched or not, goes to the "Inbox log" tab with a
 *    Gmail link, so nothing is changed silently.
 *  - Processed threads get the Gmail label "tracker-logged" and are skipped next time.
 */

const TRACKER = 'Tracker';
const LOG = 'Inbox log';
const DONE_LABEL = 'tracker-logged';
const LOOKBACK = '3d';

// Tracker columns (1-based)
const COL = { num: 1, status: 2, company: 3, role: 4, next: 10, applied: 12,
              reply: 13, followup: 17, stage: 19 };

const OPEN = /Applied|Long silence|Interviewing|Ready/i;

// Strict on purpose: confirmation emails often contain words like "unfortunately"
// or "other candidates" in their small print, so single words are not enough.
const REJECT = /unfortunately,? (we|this time|your)|regret to inform|(not|won['’]t) be (moving|proceeding|progressing)|not (moving|proceeding|progressing) (forward )?with your|decided not to (proceed|move|progress|invite|pursue)|(moving|move|proceed) forward with other candidates|not a match for|niestety.{0,80}nie|nie zaprosimy|nie będziemy kontynuować|zdecydowaliśmy.{0,40}inn|nie zakwalifikował/i;
const INVITE = /invite you|invitation to|schedule (a|an) (call|chat|meeting|interview)|your availability|book a (time|slot)|next step|zaprosić|zapraszam|zaproszenie|termin (rozmowy|spotkania)|dostępnoś/i;
const CONFIRM = /received your application|thank you for (applying|your application)|application has been (received|registered)|otrzymaliśmy|dziękujemy za (aplikacj|przesłanie|zgłoszenie)/i;

// Generic words ignored when matching company names (add your own as needed)
const STOP = new Set(['bank','group','polska','poland','systems','corporation','technology',
  'services','financial','firm','agency','logistics','fundacja','business','centrum',
  'company','international','global','solutions','team','general','consulting']);

function onOpen() {
  SpreadsheetApp.getUi().createMenu('Job tracker')
    .addItem('Add application row', 'addApplicationRow')
    .addItem('Check Gmail now', 'checkGmail')
    .addItem('Clear yellow marks', 'clearMarks')
    .addSeparator()
    .addItem('Turn on daily Gmail check', 'installDailyTrigger')
    .addToUi();
}

/** Quick add: asks for company and role, fills the rest. */
function addApplicationRow() {
  const ui = SpreadsheetApp.getUi();
  const c = ui.prompt('Company', ui.ButtonSet.OK_CANCEL);
  if (c.getSelectedButton() !== ui.Button.OK || !c.getResponseText().trim()) return;
  const r = ui.prompt('Role', ui.ButtonSet.OK_CANCEL);
  if (r.getSelectedButton() !== ui.Button.OK) return;

  const sh = SpreadsheetApp.getActive().getSheetByName(TRACKER);
  const last = lastDataRow_(sh);
  const row = last + 1;
  const prevNum = Number(sh.getRange(last, COL.num).getValue()) || last - 1;
  const company = c.getResponseText().trim();

  // Warn about repeat applications at the same company
  const companies = sh.getRange(2, COL.company, last - 1, 1).getValues().flat()
    .map(s => String(s).toLowerCase());
  const repeats = companies.filter(s => s && s.includes(company.toLowerCase())).length;
  if (repeats > 0) {
    const go = ui.alert(`You already have ${repeats} row(s) at ${company}. Add anyway?`, ui.ButtonSet.YES_NO);
    if (go !== ui.Button.YES) return;
  }

  sh.getRange(row, COL.num).setValue(prevNum + 1);
  sh.getRange(row, COL.status).setValue('🟢 Applied');
  sh.getRange(row, COL.company).setValue(company);
  sh.getRange(row, COL.role).setValue(r.getResponseText().trim());
  sh.getRange(row, COL.applied).setValue(new Date());
  sh.getRange(row, COL.stage).setValue('Applied');
  sh.setActiveRange(sh.getRange(row, 15)); // jump to Tier so you pick Tier + Channel
}

function installDailyTrigger() {
  ScriptApp.getProjectTriggers()
    .filter(t => t.getHandlerFunction() === 'checkGmail')
    .forEach(t => ScriptApp.deleteTrigger(t));
  ScriptApp.newTrigger('checkGmail').timeBased().everyDays(1).atHour(7).create();
  SpreadsheetApp.getActive().toast('Daily Gmail check is on (around 7:00).');
}

function checkGmail() {
  const ss = SpreadsheetApp.getActive();
  const sh = ss.getSheetByName(TRACKER);
  const log = ss.getSheetByName(LOG) || createLog_(ss);
  const label = GmailApp.getUserLabelByName(DONE_LABEL) || GmailApp.createLabel(DONE_LABEL);

  const last = lastDataRow_(sh);
  const data = sh.getRange(2, 1, last - 1, COL.stage).getValues();

  const q = `newer_than:${LOOKBACK} -label:${DONE_LABEL} -label:job-alerts -in:sent ` +
    '(application OR applying OR candidacy OR interview OR position OR role OR ' +
    'aplikacja OR aplikację OR rekrutacji OR rekrutacja OR stanowisko OR kandydatura)';
  const threads = GmailApp.search(q, 0, 50);
  const changes = [];

  threads.forEach(th => {
    const msg = th.getMessages().pop();            // latest message in thread
    const from = msg.getFrom();
    const subject = msg.getSubject();
    const body = msg.getPlainBody().slice(0, 3000);
    const text = `${from} ${subject} ${body}`.toLowerCase();

    // Only the subject and the start of the email count; small print is ignored.
    const head = subject + ' ' + body.slice(0, 1200);
    const isConfirmSubject = CONFIRM.test(subject) || /submitting your application|application received/i.test(subject);
    const type = REJECT.test(head) ? 'Rejection'
               : (!isConfirmSubject && INVITE.test(head)) ? 'Interview / next step'
               : CONFIRM.test(head) ? 'Confirmation'
               : null;
    if (!type) { th.addLabel(label); return; }

    const row = matchRow_(data, text);
    let action = 'Not matched: update by hand';
    const when = msg.getDate();
    const ddmm = Utilities.formatDate(when, Session.getScriptTimeZone(), 'dd.MM');

    if (row === -2) action = 'Several rows match: update by hand';
    else if (row === -3) action = 'Company found, but its rows are already closed: no change';
    else if (row > 0) {
      const r = row; // sheet row number
      const name = `${data[r - 2][2]} (row ${r})`;
      if (type === 'Rejection') {
        const before = sh.getRange(r, COL.status).getValue();
        sh.getRange(r, COL.status).setValue('🔴 Rejected')
          .setNote(`Set by Gmail script on ${ddmm}. Was: ${before}\nEmail: ${subject}`);
        sh.getRange(r, COL.reply).setValue(when);
        prependNote_(sh, r, `${ddmm} rejected (from Gmail)`);
        markChanged_(sh, r);
        action = 'Set to Rejected';
        changes.push(`Rejected: ${name}`);
      } else if (type === 'Interview / next step') {
        sh.getRange(r, COL.reply).setValue(when);
        sh.getRange(r, COL.followup).setValue(new Date());
        prependNote_(sh, r, `${ddmm} reply about next steps (from Gmail)`);
        markChanged_(sh, r);
        action = 'Reply date + follow-up set; check status';
        changes.push(`Next step: ${name}`);
      } else {
        action = 'Confirmation only, no change';
      }
    }

    log.appendRow([when, type, from, subject,
      row > 0 ? `${data[row - 2][2]}, ${data[row - 2][3]} (row ${row})` : '',
      action, `https://mail.google.com/mail/u/0/#all/${th.getId()}`]);
    th.addLabel(label);
  });

  const msgText = changes.length
    ? `Updated ${changes.length} row(s):\n${changes.join('\n')}\n\nThey are marked yellow in column A. Full list: Inbox log tab.`
    : `${threads.length} email(s) checked, no rows changed. Details: Inbox log tab.`;
  try { SpreadsheetApp.getUi().alert('Gmail check', msgText, SpreadsheetApp.getUi().ButtonSet.OK); }
  catch (e) { /* daily trigger runs without a screen: the yellow marks and Inbox log show the changes */ }
}

/** Marks a changed row: yellow # cell. Run "Clear yellow marks" once you've checked them. */
function markChanged_(sh, row) {
  sh.getRange(row, COL.num).setBackground('#fff2a8');
}

function clearMarks() {
  const sh = SpreadsheetApp.getActive().getSheetByName(TRACKER);
  sh.getRange(2, COL.num, sh.getMaxRows() - 1, 1).setBackground(null);
}

/** Returns sheet row number, -1 for no match, -2 for ambiguous, -3 if only closed rows match. */
function matchRow_(data, text) {
  const hits = [];
  let closedHit = false;
  data.forEach((r, i) => {
    const tokens = String(r[COL.company - 1]).toLowerCase()
      .replace(/[(),.]/g, ' ').split(/\s+/)
      .filter(t => t.length >= 3 && !STOP.has(t));
    if (!tokens.some(t => text.includes(t))) return;
    if (OPEN.test(String(r[COL.status - 1]))) hits.push(i); else closedHit = true;
  });
  if (hits.length === 0) return closedHit ? -3 : -1;
  if (hits.length === 1) return hits[0] + 2;

  // Several open rows at one company: pick the one whose role words fit best
  const scored = hits.map(i => {
    const words = String(data[i][COL.role - 1]).toLowerCase().split(/[^a-ząćęłńóśźż]+/)
      .filter(w => w.length >= 4);
    return { i, s: words.filter(w => text.includes(w)).length };
  }).sort((a, b) => b.s - a.s);
  return scored[0].s > scored[1].s ? scored[0].i + 2 : -2;
}

function prependNote_(sh, row, note) {
  const cell = sh.getRange(row, COL.next);
  const old = String(cell.getValue() || '');
  cell.setValue(old ? `${note}\n${old}` : note);
}

function lastDataRow_(sh) {
  const vals = sh.getRange(1, COL.company, sh.getLastRow(), 1).getValues().flat();
  for (let i = vals.length - 1; i >= 0; i--) if (String(vals[i]).trim()) return i + 1;
  return 1;
}

function createLog_(ss) {
  const s = ss.insertSheet(LOG);
  s.appendRow(['Received', 'Type', 'From', 'Subject', 'Matched row', 'What the script did', 'Gmail link']);
  s.getRange('A1:G1').setFontWeight('bold').setBackground('#1f3864').setFontColor('#ffffff');
  s.setFrozenRows(1);
  s.getRange('A:A').setNumberFormat('yyyy-mm-dd');
  return s;
}
