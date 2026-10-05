/*************************************************************
 * MENTOR ALLOCATION & ATTENDANCE FOLLOW-UP  — BACKEND  (v2)
 *************************************************************/

const MENTOR_SHEET_ID   = '10HJebqzvQvCnuIFCKu88zJ5E-DYf9qqIWq686NXXBbk';
const MASTER_SHEET_NAME = 'Mentor Name';
const ALLOCATION_SHEETS = ['UG 2', 'UG 3', 'UG 5', 'UG 7', 'IMCA 7', 'PG 3'];
const FOLLOWUP_SHEET    = 'FollowUpRemarks';
const UPLOADED_SHEET    = 'UploadedAttendance';

function doGet() {
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle('Mentor Allocation & Attendance Follow-Up')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function getSS_() { return SpreadsheetApp.openById(MENTOR_SHEET_ID); }

function normEnr_(v) {
  if (v === null || v === undefined) return '';
  let s = String(v).trim();
  s = s.replace(/\.0+$/, '');
  s = s.replace(/[\s\u00A0]+/g, '');
  return s.toUpperCase();
}
function normName_(v) {
  if (v === null || v === undefined) return '';
  return String(v).toUpperCase().replace(/\s+/g, ' ').trim();
}
function findHeader_(rows) {
  const limit = Math.min(20, rows.length);
  for (let i = 0; i < limit; i++) {
    const line = rows[i].map(c => String(c || '').toLowerCase()).join('|');
    if (line.includes('enroll') && line.includes('mentor')) {
      const headers = rows[i].map(c => String(c || '').toLowerCase());
      const col = {
        enrollment: headers.findIndex(h => h.includes('enroll')),
        name:       headers.findIndex(h =>
                       (h.includes('student') && h.includes('name')) ||
                       h === 'student' || h === 'name' ||
                       h.includes('name as per document') ||
                       h === 'studentname'),
        mentor:     headers.findIndex(h => h.includes('mentor')),
        course:     headers.findIndex(h => h.includes('course') || h.includes('program'))
      };
      if (col.enrollment !== -1 && col.mentor !== -1) return { row: i, col };
    }
  }
  return null;
}

function getMentorAllocation() {
  try {
    const ss = getSS_();
    const map = {};

    ALLOCATION_SHEETS.forEach(name => {
      const sheet = ss.getSheetByName(name);
      if (!sheet) return;
      const values  = sheet.getDataRange().getValues();
      const display = sheet.getDataRange().getDisplayValues();
      const header  = findHeader_(values);
      if (!header) return;
      const { row, col } = header;
      for (let i = row + 1; i < values.length; i++) {
        const enroll = normEnr_(display[i][col.enrollment]);
        if (!enroll) continue;
        const mentor = String(display[i][col.mentor] || '').trim();
        if (!mentor || mentor.startsWith('=') || mentor.startsWith('#')) continue;
        const name   = col.name   !== -1 ? String(display[i][col.name]   || '').trim() : '';
        const course = col.course !== -1 ? String(display[i][col.course] || '').trim() : '';
        if (!map[enroll]) {
          map[enroll] = { mentor, studentName: name, course, source: name };
        }
      }
    });

    const master = ss.getSheetByName(MASTER_SHEET_NAME);
    if (master) {
      const values  = master.getDataRange().getValues();
      const display = master.getDataRange().getDisplayValues();
      const header  = findHeader_(values);
      if (header) {
        const { row, col } = header;
        for (let i = row + 1; i < values.length; i++) {
          const enroll = normEnr_(display[i][col.enrollment]);
          if (!enroll || map[enroll]) continue;
          const mentor = String(display[i][col.mentor] || '').trim();
          if (!mentor || mentor.startsWith('=') || mentor.startsWith('#')) continue;
          const name   = col.name   !== -1 ? String(display[i][col.name]   || '').trim() : '';
          const course = col.course !== -1 ? String(display[i][col.course] || '').trim() : '';
          map[enroll] = { mentor, studentName: name, course, source: MASTER_SHEET_NAME };
        }
      }
    }

    const data = Object.keys(map).map(k => ({
      enrollment: k,
      mentor: map[k].mentor,
      studentName: map[k].studentName,
      course: map[k].course,
      source: map[k].source
    }));

    return { success: true, data, count: data.length };
  } catch (e) { return { success: false, error: e.message }; }
}

function getOrCreate_(name, headers) {
  const ss = getSS_();
  let sh = ss.getSheetByName(name);
  if (!sh) {
    sh = ss.insertSheet(name);
    sh.appendRow(headers);
    sh.setFrozenRows(1);
    sh.getRange(1,1,1,headers.length).setFontWeight('bold');
  }
  return sh;
}
function getFollowUpSheet_() {
  return getOrCreate_(FOLLOWUP_SHEET, [
    'UpdatedAt','Enrollment No','Student Name','Mentor','Course','Semester','Division',
    'This Week %','Last Week %','Difference','Below Threshold','Follow-Up Remark','Updated By'
  ]);
}
function getUploadedSheet_() {
  return getOrCreate_(UPLOADED_SHEET, [
    'UploadedAt','Week','FileName','Enrollment No','Student Name','Roll No',
    'Guardian Phone','CH','CA','P (%)','Division','Semester','Course'
  ]);
}

function saveBulkFollowUp(rows) {
  try {
    const sh = getFollowUpSheet_();
    const data = sh.getDataRange().getValues();
    const headers = data[0];
    const enrCol  = headers.indexOf('Enrollment No');

    const existing = {};
    for (let i = 1; i < data.length; i++) {
      existing[normEnr_(data[i][enrCol])] = i + 1;
    }

    const now = new Date();
    const user = Session.getActiveUser().getEmail() || 'unknown';
    const toAppend = [];

    rows.forEach(r => {
      const row = [
        now,
        r.enrollment, r.studentName || '', r.mentor || '',
        r.course || '', r.sem || '', r.division || '',
        r.thisWeekP || 0, r.lastWeekP || 0, r.diff || 0,
        r.thresholdFlag ? 'Yes' : 'No',
        r.followUp || '', user
      ];
      const key = normEnr_(r.enrollment);
      if (existing[key]) {
        sh.getRange(existing[key], 1, 1, row.length).setValues([row]);
      } else {
        toAppend.push(row);
      }
    });
    if (toAppend.length) {
      sh.getRange(sh.getLastRow() + 1, 1, toAppend.length, toAppend[0].length)
        .setValues(toAppend);
    }
    return { success: true, saved: rows.length };
  } catch (e) { return { success: false, error: e.message }; }
}

function getSavedFollowUps() {
  try {
    const sh = getFollowUpSheet_();
    const data = sh.getDataRange().getValues();
    if (data.length < 2) return { success: true, data: {} };
    const headers = data[0];
    const enrCol  = headers.indexOf('Enrollment No');
    const remCol  = headers.indexOf('Follow-Up Remark');
    const map = {};
    for (let i = 1; i < data.length; i++) {
      map[normEnr_(data[i][enrCol])] = data[i][remCol] || '';
    }
    return { success: true, data: map };
  } catch (e) { return { success: false, error: e.message }; }
}

function saveUploadedData(week, students, fileName) {
  try {
    const sh = getUploadedSheet_();
    const data = sh.getDataRange().getValues();
    const h = data[0];
    const idx = {
      week: h.indexOf('Week'),
      enr:  h.indexOf('Enrollment No')
    };

    const existing = {};
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][idx.week]).toLowerCase() !== week) continue;
      const key = normEnr_(data[i][idx.enr]);
      if (key) existing[key] = i + 1;
    }

    const now = new Date();
    const toAppend = [];

    students.forEach(s => {
      const row = [
        now, week, fileName || '',
        s.enrollment || '', s.name || '', s.roll || '',
        s.guardian || '', s.ch || 0, s.ca || 0, s.p || 0,
        s.division || '', s.sem || '', s.course || ''
      ];
      const key = normEnr_(s.enrollment);
      if (!key) return;
      if (existing[key]) {
        sh.getRange(existing[key], 1, 1, row.length).setValues([row]);
      } else {
        toAppend.push(row);
      }
    });

    if (toAppend.length) {
      sh.getRange(sh.getLastRow() + 1, 1, toAppend.length, toAppend[0].length)
        .setValues(toAppend);
    }
    return { success: true, added: toAppend.length, total: students.length };
  } catch (e) { return { success: false, error: e.message }; }
}

function getUploadedData() {
  try {
    const sh = getUploadedSheet_();
    const data = sh.getDataRange().getValues();
    if (data.length < 2) return { success: true, thisWeek: [], lastWeek: [] };
    const h = data[0];
    const idx = {
      week: h.indexOf('Week'), fileName: h.indexOf('FileName'),
      enr: h.indexOf('Enrollment No'), name: h.indexOf('Student Name'),
      roll: h.indexOf('Roll No'), guardian: h.indexOf('Guardian Phone'),
      ch: h.indexOf('CH'), ca: h.indexOf('CA'), p: h.indexOf('P (%)'),
      div: h.indexOf('Division'), sem: h.indexOf('Semester'), course: h.indexOf('Course')
    };
    const thisWeek = [], lastWeek = [];
    for (let i = 1; i < data.length; i++) {
      const r = data[i];
      const obj = {
        enrollment: normEnr_(r[idx.enr]),
        name: String(r[idx.name] || '').trim(),
        roll: String(r[idx.roll] || '').trim(),
        guardian: String(r[idx.guardian] || '').trim(),
        ch: Number(r[idx.ch]) || 0,
        ca: Number(r[idx.ca]) || 0,
        p: Number(r[idx.p]) || 0,
        division: String(r[idx.div] || '').trim(),
        sem: String(r[idx.sem] || '').trim(),
        course: String(r[idx.course] || '').trim()
      };
      if (String(r[idx.week]).toLowerCase() === 'this') thisWeek.push(obj);
      else if (String(r[idx.week]).toLowerCase() === 'last') lastWeek.push(obj);
    }
    return { success: true, thisWeek, lastWeek };
  } catch (e) { return { success: false, error: e.message }; }
}

function clearUploadedData(week) {
  try {
    const sh = getUploadedSheet_();
    const data = sh.getDataRange().getValues();
    const weekCol = data[0].indexOf('Week');
    for (let i = data.length - 1; i >= 1; i--) {
      if (String(data[i][weekCol]).toLowerCase() === week) sh.deleteRow(i + 1);
    }
    return { success: true };
  } catch (e) { return { success: false, error: e.message }; }
}

function exportFollowUpToSheet(rows) {
  try {
    const ss = getSS_();
    const name = 'FollowUp_' + Utilities.formatDate(
      new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd_HH-mm');
    const sh = ss.insertSheet(name);
    const headers = ['Enrollment No','Student Name','Mentor','Course','Semester',
                     'Division','This Week %','Last Week %','Difference',
                     'Below Threshold','Follow-Up Remark'];
    sh.appendRow(headers);
    sh.getRange(1,1,1,headers.length).setFontWeight('bold');
    sh.setFrozenRows(1);
    if (rows.length) {
      const out = rows.map(d => ([
        d.enrollment, d.studentName, d.mentor, d.course, d.sem, d.division,
        d.thisWeekP, d.lastWeekP, d.diff,
        d.thresholdFlag ? 'Yes' : 'No', d.followUp || ''
      ]));
      sh.getRange(2,1,out.length,headers.length).setValues(out);
    }
    sh.autoResizeColumns(1, headers.length);
    return { success: true, url: ss.getUrl() + '#gid=' + sh.getSheetId(), name };
  } catch (e) { return { success: false, error: e.message }; }
}

function testMentors() {
  const r = getMentorAllocation();
  Logger.log(r.success ? ('count=' + r.count + ' sample=' +
      JSON.stringify(r.data.slice(0,3))) : r.error);
}
