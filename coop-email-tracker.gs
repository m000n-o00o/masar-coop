/**
 * Co-op Email Tracker (Google Apps Script)
 * Scans your Gmail for co-op / internship emails, guesses the status
 * (offer, rejection, interview, received), and logs each one to this Google Sheet.
 * Everything runs inside your own Google account. Nothing is sent anywhere else.
 *
 * Setup:
 * 1. Create a new Google Sheet.
 * 2. Extensions > Apps Script, delete the sample code, paste this file, save.
 * 3. Choose the "setup" function and press Run, then approve the permissions.
 *    It scans the last DAYS_BACK days and installs a daily automatic scan.
 */

const CONFIG = {
  SHEET_NAME: 'Applications',
  DAYS_BACK: 180,            // how far back the first scan looks
  RUN_HOUR: 7,               // daily scan hour (your account's time zone)
  LABEL: 'Co-op',            // Gmail label added to matched threads ('' to disable)
  // Gmail search. Edit freely: add company names or remove words that bring noise.
  QUERY: '(internship OR intern OR "co-op" OR coop OR "cooperative training" OR trainee OR ' +
         '"تدريب تعاوني" OR "التدريب التعاوني" OR "تدريب صيفي" OR "برنامج تدريب") ' +
         '-category:promotions -category:social -in:sent',
};

// Checked in this order: the first match decides the status.
const RULES = [
  { status: 'قبول', words: ['congratulations', 'pleased to offer', 'offer letter', 'we are pleased to inform',
      'you have been selected', 'you have been accepted', 'مبروك', 'تهانينا', 'تم قبولك', 'يسعدنا إبلاغك', 'خطاب القبول'] },
  { status: 'رفض', words: ['unfortunately', 'regret to inform', 'not been selected', 'not to move forward',
      'other candidates', 'will not be progressing', 'نأسف', 'نعتذر', 'لم يتم اختيارك', 'لم يتم قبول', 'عدم قبول'] },
  { status: 'مقابلة', words: ['invite you to an interview', 'invitation to interview', 'interview invitation',
      'schedule your interview', 'schedule an interview', 'online assessment', 'hackerrank', 'codility', 'shortlisted',
      'دعوة لمقابلة', 'موعد المقابلة', 'موعد مقابلة', 'اختبار إلكتروني', 'تم ترشيحك'] },
  { status: 'تم الاستلام', words: ['received your application', 'thank you for applying', 'thank you for your application',
      'application has been submitted', 'تم استلام', 'شكرا لتقديمك', 'شكراً لتقديمك', 'تم تقديم طلبك'] },
  // Anything else that still looks application-related, so nothing slips through unlogged.
  { status: 'تحديث', words: ['application', 'interview', 'candidate', 'طلبك', 'مقابلة', 'المتقدم'] },
];

const HEADERS = ['Date', 'Company', 'Status', 'Subject', 'From', 'Link', 'MessageId'];

function setup() {
  getSheet_();
  ScriptApp.getProjectTriggers()
    .filter(t => t.getHandlerFunction() === 'scanInbox')
    .forEach(t => ScriptApp.deleteTrigger(t));
  ScriptApp.newTrigger('scanInbox').timeBased().everyDays(1).atHour(CONFIG.RUN_HOUR).create();
  scanInbox();
}

function scanInbox() {
  const sheet = getSheet_();
  const seen = new Set(
    sheet.getLastRow() > 1
      ? sheet.getRange(2, 7, sheet.getLastRow() - 1, 1).getValues().map(r => String(r[0]))
      : []
  );
  const label = CONFIG.LABEL ? (GmailApp.getUserLabelByName(CONFIG.LABEL) || GmailApp.createLabel(CONFIG.LABEL)) : null;
  const query = CONFIG.QUERY + ' newer_than:' + CONFIG.DAYS_BACK + 'd';
  const rows = [];

  for (let start = 0; start < 500; start += 100) {
    const threads = GmailApp.search(query, start, 100);
    if (!threads.length) break;
    threads.forEach(thread => {
      thread.getMessages().forEach(msg => {
        const id = msg.getId();
        if (seen.has(id)) return;
        const subject = msg.getSubject() || '';
        const text = (subject + ' ' + msg.getPlainBody().slice(0, 4000)).toLowerCase();
        const status = classify_(text);
        if (!status) return;   // matched the search but looks unrelated to an application
        rows.push([msg.getDate(), company_(msg.getFrom()), status, subject, msg.getFrom(),
                   'https://mail.google.com/mail/u/0/#all/' + thread.getId(), id]);
        seen.add(id);
      });
      if (label) thread.addLabel(label);
    });
  }

  if (rows.length) {
    rows.sort((a, b) => a[0] - b[0]);
    sheet.getRange(sheet.getLastRow() + 1, 1, rows.length, HEADERS.length).setValues(rows);
  }
}

function classify_(text) {
  for (const rule of RULES) {
    if (rule.words.some(w => text.indexOf(w.toLowerCase()) !== -1)) return rule.status;
  }
  return '';
}

// "Talent Team at Acme <careers@acme.com>" -> "Talent Team at Acme";
// falls back to the sender's domain when there is no display name.
function company_(from) {
  const m = from.match(/^\s*"?([^"<]*?)"?\s*<([^>]+)>/);
  const name = m ? m[1].trim() : '';
  const email = m ? m[2] : from;
  const domain = (email.split('@')[1] || '').replace(/^(mail|email|jobs|careers|hr|no-?reply)\./, '');
  const generic = /(gmail|outlook|hotmail|yahoo|icloud|myworkday|workday|successfactors|greenhouse|lever|taleo|smartrecruiters)\./i;
  if (name) return name;
  return generic.test(domain) ? email : domain.split('.')[0];
}

function getSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(CONFIG.SHEET_NAME) || ss.insertSheet(CONFIG.SHEET_NAME);
  if (sheet.getLastRow() === 0) {
    sheet.appendRow(HEADERS);
    sheet.setFrozenRows(1);
    sheet.getRange('A:A').setNumberFormat('yyyy-mm-dd');
  }
  return sheet;
}
