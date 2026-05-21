/** ========= GITHUB PAGES API MODE =========
 * วิธีใช้:
 * 1) นำไฟล์นี้ไปวางใน Google Apps Script ที่ผูกกับ Google Sheet เดิม
 * 2) Deploy > New deployment > Web app
 *    - Execute as: Me
 *    - Who has access: Anyone / Anyone with the link
 * 3) คัดลอก Web app URL ไปใส่ใน index.html ตรงตัวแปร API_ENDPOINT
 */

const API_ACTIONS = {
  saveAttendance,
  getReport,
  exportReportCSV,
  getEditRecords,
  updateRecord,
  getStudents,
  getSubmissionStatus,
  getWeeklySubmissionStatus,
  checkAttendanceExists,
  getMissingAttendanceDates
};

function doGet(e) {

  // ===== API MODE =====
  if (e && e.parameter.action) {

    const action = e.parameter.action;

    // โหลดนักเรียน
    if (action === 'getStudents') {

      const room = e.parameter.room || '';

      return ContentService
        .createTextOutput(JSON.stringify({
          ok: true,
          result: getStudents(room)
        }))
        .setMimeType(ContentService.MimeType.JSON);
    }

    // โหลดรายงาน
    if (action === 'getReport') {

      return ContentService
        .createTextOutput(JSON.stringify(
          getReport(
            e.parameter.room || '',
            e.parameter.start || '',
            e.parameter.end || ''
          )
        ))
        .setMimeType(ContentService.MimeType.JSON);
    }

    // โหลดสถานะ
    if (action === 'getSubmissionStatus') {

      return ContentService
        .createTextOutput(JSON.stringify(
          getSubmissionStatus(e.parameter.date || '')
        ))
        .setMimeType(ContentService.MimeType.JSON);
    }
  }

  // ===== WEB MODE =====
  return HtmlService
    .createHtmlOutputFromFile('index');
}

function doPost(e) {
  let body = {};

  try {
    body = JSON.parse((e && e.postData && e.postData.contents) || '{}');
  } catch (err) {
    body = {};
  }

  return handleApi_(body, 'POST');
}

function handleApi_(req, method) {
  const callback = cleanCallback_(req.callback || '');
  const action = String(req.action || '').trim();

  let payload;

  try {
    if (!API_ACTIONS[action]) {
      throw new Error('Unknown action: ' + action);
    }

    let args = req.args;

    if (typeof args === 'string') {
      args = JSON.parse(args || '[]');
    }

    if (!Array.isArray(args)) args = [];

    const result = API_ACTIONS[action].apply(null, args);

    payload = {
      ok: true,
      result: result
    };

  } catch (err) {
    payload = {
      ok: false,
      error: String(err && err.message ? err.message : err)
    };
  }

  const json = JSON.stringify(payload);
  const output = callback ? `${callback}(${json});` : json;
  const mime = callback ? ContentService.MimeType.JAVASCRIPT : ContentService.MimeType.JSON;

  return ContentService.createTextOutput(output).setMimeType(mime);
}

function cleanCallback_(s) {
  s = String(s || '').trim();
  return /^[A-Za-z_$][0-9A-Za-z_$\.]*$/.test(s) ? s : '';
}

/** ========= CONFIG ========= **/
const ROOMS = ['ม.4/1','ม.4/2','ม.4/3','ม.4/4','ม.4/5','ม.4/6','ม.4/7','ม.4/8','ม.4/9','ม.4/10','ม.4/11'];
const HEADER_STD = ['date','room','student_no','student_name','status','timestamp'];

/** ========= UTIL ========= **/
const TZ = () => (Session.getScriptTimeZone() || 'Asia/Bangkok');

function fmtISO(d){
  return Utilities.formatDate(d, TZ(), 'yyyy-MM-dd');
}

function nowTS(){
  return Utilities.formatDate(new Date(), TZ(), 'yyyy-MM-dd HH:mm:ss');
}

function toISO(value){
  if (value instanceof Date) return fmtISO(value);

  if (typeof value === 'number') {
    const d = new Date(Math.round((value - 25569) * 86400 * 1000));
    return fmtISO(d);
  }

  const s = String(value || '').trim();
  if (!s) return '';

  if (/^\d{4}-\d{2}-\d{2}$/.test(s)) return s;

  const m = s.match(/^(\d{1,2})[\/\-](\d{1,2})[\/\-](\d{4})$/);
  if (m) {
    const dd = ('0' + m[1]).slice(-2);
    const mm = ('0' + m[2]).slice(-2);
    const yy = m[3];
    return `${yy}-${mm}-${dd}`;
  }

  const d = new Date(s);
  if (!isNaN(d)) return fmtISO(d);

  return s;
}

function findHeaderIndexMap_(headers){
  const norm = v => String(v || '').toLowerCase().trim();

  const want = {
    date:['date','วันที่','วัน'],
    room:['room','ห้อง','ชั้น'],
    student_no:['student_no','no','เลขที่','เลข'],
    student_name:['student_name','name','ชื่อ','ชื่อ-สกุล','ชื่อสกุล'],
    status:['status','สถานะ','เช็คชื่อ','ผล'],
    timestamp:['timestamp','เวลา','บันทึกเมื่อ']
  };

  const map = {};
  const H = headers.map(norm);

  Object.keys(want).forEach(k => {
    let idx = -1;

    for (let i = 0; i < H.length; i++) {
      if (want[k].some(w => H[i] === w)) {
        idx = i;
        break;
      }
    }

    if (idx < 0) {
      if (k === 'date') idx = 0;
      else if (k === 'room') idx = 1;
      else if (k === 'student_no') idx = 2;
      else if (k === 'student_name') idx = 3;
      else if (k === 'status') idx = 4;
      else if (k === 'timestamp') idx = 5;
    }

    map[k] = idx;
  });

  return map;
}

function isAttendanceSheetName_(name){
  if (name === 'Attendance') return true;
  if (ROOMS.includes(name)) return true;
  return ROOMS.some(r => name === `Attendance ${r}`);
}

function ensureStdHeader_(sh){
  const width = HEADER_STD.length;
  const first = sh.getRange(1, 1, 1, width).getValues()[0];

  if (first.join(',') !== HEADER_STD.join(',')) {
    sh.getRange(1, 1, 1, width).setValues([HEADER_STD]);
  }
}

function readSheetObjects_(sh){
  const vals = sh.getDataRange().getValues();
  if (vals.length <= 1) return [];

  const header = vals[0];
  const idx = findHeaderIndexMap_(header);
  const out = [];

  for (let r = 1; r < vals.length; r++) {
    const row = vals[r];

    const date = toISO(row[idx.date]);
    const room = String(row[idx.room] || '').trim();
    const no = Number(row[idx.student_no] || 0);
    const name = String(row[idx.student_name] || '').trim();
    const status = String(row[idx.status] || '').trim();
    const ts = row[idx.timestamp] ? String(row[idx.timestamp]) : '';

    if (!date && !room && !no && !name && !status) continue;

    out.push({
      __sheet: sh.getName(),
      __row: r + 1,
      date: date,
      room: room,
      student_no: no,
      student_name: name,
      status: status,
      timestamp: ts
    });
  }

  return out;
}

function readAllAttendance_(){
  const ss = SpreadsheetApp.getActive();
  const result = [];

  ss.getSheets().forEach(sh => {
    const nm = sh.getName();

    if (isAttendanceSheetName_(nm)) {
      result.push.apply(result, readSheetObjects_(sh));
    }
  });

  return result;
}

function sheetForRoom_(room){
  const ss = SpreadsheetApp.getActive();

  const legacy = ss.getSheetByName('Attendance');
  if (legacy) {
    ensureStdHeader_(legacy);
    return legacy;
  }

  let sh = ss.getSheetByName(room) || ss.getSheetByName(`Attendance ${room}`);

  if (!sh) {
    sh = ss.insertSheet(`Attendance ${room}`);
  }

  ensureStdHeader_(sh);
  return sh;
}

/** ========= BACKEND API ========= **/

function saveAttendance(dateISO, room, entries){
  try {
    const sh = sheetForRoom_(room);
    const dateStr = toISO(dateISO);

    const all = readSheetObjects_(sh);
    const keep = all.filter(r => !(r.date === dateStr && r.room === room));

    sh.clearContents();
    ensureStdHeader_(sh);

    if (keep.length) {
      const vals = keep.map(r => [
        r.date,
        r.room,
        r.student_no,
        r.student_name,
        r.status,
        r.timestamp || ''
      ]);

      sh.getRange(2, 1, vals.length, HEADER_STD.length).setValues(vals);
    }

    const add = entries.map(e => [
      dateStr,
      room,
      e.no,
      e.name,
      normalizeStatus_(e.status),
      nowTS()
    ]);

    if (add.length) {
      sh.getRange(sh.getLastRow() + 1, 1, add.length, HEADER_STD.length).setValues(add);
    }

    return {ok:true, saved:add.length};

  } catch(e) {
    return {ok:false, error:String(e)};
  }
}

function getReport(roomOrEmpty, startISO, endISO){
  try {
    const start = toISO(startISO);
    const end = toISO(endISO);

    const rows = readAllAttendance_()
      .filter(r => {
        const d = toISO(r.date);
        const okDate = d >= start && d <= end;
        const okRoom = !roomOrEmpty || r.room === roomOrEmpty;
        return okDate && okRoom && r.status;
      })
      .sort((a,b) =>
        a.date.localeCompare(b.date) ||
        a.room.localeCompare(b.room) ||
        a.student_no - b.student_no
      );

    return {ok:true, rows:rows};

  } catch(e) {
    return {ok:false, error:String(e), rows:[]};
  }
}

function exportReportCSV(roomOrEmpty, startISO, endISO){
  const rep = getReport(roomOrEmpty, startISO, endISO);

  if (!rep.ok || !rep.rows.length) return '';

  const dateSet = new Set(rep.rows.map(r => r.date));
  const dates = Array.from(dateSet).sort();

  const byStu = {};

  rep.rows.forEach(r => {
    const key = `${r.room}|${r.student_no}|${r.student_name}`;

    if (!byStu[key]) {
      byStu[key] = {
        room:r.room,
        no:r.student_no,
        name:r.student_name,
        att:{},
        M:0,
        K:0,
        L:0
      };
    }

    const s = normalizeStatus_(r.status);
    byStu[key].att[r.date] = s;

    if (s === 'M') byStu[key].M++;
    else if (s === 'K') byStu[key].K++;
    else if (s === 'L') byStu[key].L++;
  });

  const tz = TZ();

  const header = [
    'ห้อง',
    'เลขที่',
    'ชื่อ-สกุล',
    ...dates.map(d => Utilities.formatDate(new Date(d + 'T00:00:00'), tz, 'dd/MM/yyyy')),
    'มา(ครั้ง)',
    'ขาด(ครั้ง)',
    'ลา(ครั้ง)',
    'รวมทั้งหมด'
  ];

  const text = {
    M:'มา',
    K:'ขาด',
    L:'ลา'
  };

  const body = Object.values(byStu)
    .sort((a,b) => a.room.localeCompare(b.room) || a.no - b.no)
    .map(st => {
      const daily = dates.map(d => text[st.att[d]] || '');
      const total = st.M + st.K + st.L;

      return [
        st.room,
        st.no,
        st.name,
        ...daily,
        st.M,
        st.K,
        st.L,
        total
      ];
    });

  const csv = [header, ...body].map(row =>
    row.map(v => {
      const s = String(v ?? '');
      return /[",\n]/.test(s) ? `"${s.replace(/"/g,'""')}"` : s;
    }).join(',')
  ).join('\n');

  return csv;
}

function getEditRecords(room, startISO, endISO){
  try {
    const start = toISO(startISO);
    const end = toISO(endISO);

    const rows = readAllAttendance_()
      .filter(r => {
        return r.room === room &&
          toISO(r.date) >= start &&
          toISO(r.date) <= end &&
          r.status;
      })
      .sort((a,b) => a.date.localeCompare(b.date) || a.student_no - b.student_no);

    return {ok:true, rows:rows};

  } catch(e) {
    return {ok:false, error:String(e), rows:[]};
  }
}

function updateRecord(sheetName, rowId, newStatus){
  try {
    const ss = SpreadsheetApp.getActive();
    const sh = ss.getSheetByName(sheetName);

    if (!sh) return {ok:false, error:'sheet not found'};

    ensureStdHeader_(sh);

    const idx = findHeaderIndexMap_(sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0]);
    const r = Number(rowId);

    if (r < 2 || r > sh.getLastRow()) {
      return {ok:false, error:'invalid row'};
    }

    sh.getRange(r, idx.status + 1).setValue(normalizeStatus_(newStatus));
    sh.getRange(r, idx.timestamp + 1).setValue(nowTS());

    return {ok:true};

  } catch(e) {
    return {ok:false, error:String(e)};
  }
}

function getStudents(room){
  const ss = SpreadsheetApp.getActive();
  const st = ss.getSheetByName('Students');

  if (!st) {
    return Array.from({length:35}, (_,i) => ({
      no:i + 1,
      name:`นักเรียน ${room.replace('ม.4/','')}/${i + 1}`
    }));
  }

  const vals = st.getDataRange().getValues();
  if (!vals.length) return [];

  const h = vals[0].map(x => String(x).trim());
  const iRoom = h.indexOf('room');
  const iNo = h.indexOf('no');
  const iName = h.indexOf('name');

  if (iRoom < 0 || iNo < 0 || iName < 0) {
    return Array.from({length:35}, (_,i) => ({
      no:i + 1,
      name:`นักเรียน ${room.replace('ม.4/','')}/${i + 1}`
    }));
  }

  const out = [];

  for (let i = 1; i < vals.length; i++) {
    if (String(vals[i][iRoom]) === room) {
      out.push({
        no:Number(vals[i][iNo]),
        name:String(vals[i][iName])
      });
    }
  }

  out.sort((a,b) => a.no - b.no);
  return out;
}

function getSubmissionStatus(dateISO) {
  const d = toISO(dateISO || new Date());
  const all = readAllAttendance_().filter(r => r.date === d);

  const seen = {};

  all.forEach(r => {
    seen[r.room] = (seen[r.room] || 0) + 1;
  });

  const rows = ROOMS.map(room => ({
    room:room,
    submitted:!!seen[room],
    count:seen[room] || 0
  }));

  return {
    ok:true,
    rows:rows,
    date:d
  };
}
function getWeeklySubmissionStatus(weekDates) {
  try {
    const all = readAllAttendance_();

    const result = ROOMS.map(room => {
      const days = {};

      weekDates.forEach(dateISO => {
        const count = all.filter(r => r.room === room && r.date === dateISO).length;

        days[dateISO] = {
          submitted: count > 0,
          count: count
        };
      });

      return {
        room: room,
        days: days
      };
    });

    return {
      ok: true,
      rows: result
    };

  } catch(e) {
    return {
      ok: false,
      error: String(e),
      rows: []
    };
  }
}

/** ========= HELPERS ========= **/

function normalizeStatus_(s){
  const x = String(s || '').trim().toUpperCase();

  if (x === 'มา' || x === 'P' || x === 'M') return 'M';
  if (x === 'ขาด' || x === 'A' || x === 'K') return 'K';
  if (x === 'ลา' || x === 'L') return 'L';

  return x || 'M';
}
function checkAttendanceExists(dateISO, room) {
  try {
    const dateStr = toISO(dateISO);
    const all = readAllAttendance_();

    const found = all.some(r => {
      return r.date === dateStr && r.room === room && r.status;
    });

    return {
      ok: true,
      exists: found
    };

  } catch(e) {
    return {
      ok: false,
      exists: false,
      error: String(e)
    };
  }
}

function getMissingAttendanceDates(room, startISO, endISO) {
  try {
    const start = toISO(startISO);
    const end = toISO(endISO);

    const all = readAllAttendance_();
    const checkedDates = {};

    all.forEach(r => {
      if (r.room === room && r.status) {
        checkedDates[r.date] = true;
      }
    });

    const missingDates = [];
    let current = new Date(start + 'T00:00:00');
    const last = new Date(end + 'T00:00:00');

    while (current <= last) {
      const day = current.getDay();

      if (day >= 1 && day <= 5) {
        const dateISO = Utilities.formatDate(current, TZ(), 'yyyy-MM-dd');

        if (!checkedDates[dateISO]) {
          missingDates.push(dateISO);
        }
      }

      current.setDate(current.getDate() + 1);
    }

    return {
      ok: true,
      missingDates: missingDates
    };

  } catch(e) {
    return {
      ok: false,
      error: String(e),
      missingDates: []
    };
  }
}

