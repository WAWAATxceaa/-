// =========================================================
// ระบบปฏิทินการปฏิบัติงาน - Google Apps Script Backend (Real-time Sync)
// Code.gs
// =========================================================

// ชื่อ Sheet ใน Google Spreadsheet
var SHEET_EVENTS   = 'Events';
var SHEET_USERS    = 'Users';
var SHEET_HOLIDAYS = 'Holidays';

// กำหนดคอลัมน์ตามข้อกำหนดล่าสุด (เพิ่ม updatedAt)
var EVENT_COLS   = ['id','title','type','status','assignee','start','end','startTime','endTime','location','detail','updatedAt'];
var USER_COLS    = ['email','name','department','role'];
var HOLIDAY_COLS = ['date','name','type'];

// =========================================================
// HTTP Handlers (Web App API & Web App HTML Service)
// =========================================================

function doGet(e) {
  // หากไม่มี action parameter ให้แสดงหน้าเว็บ HTML (กรณี Deploy และเปิดผ่าน Web App URL โดยตรง)
  if (!e || !e.parameter || !e.parameter.action) {
    try {
      var html = HtmlService.createTemplateFromFile('index').evaluate();
      html.setTitle('ระบบปฏิทินการปฏิบัติงาน');
      html.addMetaTag('viewport', 'width=device-width, initial-scale=1.0');
      html.setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
      return html;
    } catch (err) {
      return buildResponse({
        success: true,
        message: 'ระบบปฏิทินการปฏิบัติงาน API พร้อมใช้งานสำหรับ Firestore Sync',
        version: '2.0-realtime'
      });
    }
  }

  try {
    var action = e.parameter.action || '';
    var result = routeAction(action, e.parameter, 'GET');
    return buildResponse(result);
  } catch (err) {
    return buildResponse({ success: false, error: err.message });
  }
}

function doPost(e) {
  try {
    var payload = {};
    if (e && e.postData && e.postData.contents) {
      try {
        payload = JSON.parse(e.postData.contents);
      } catch (jsonErr) {
        payload = e.parameter || {};
      }
    } else if (e && e.parameter) {
      payload = e.parameter;
    }

    var action = payload.action || (e && e.parameter && e.parameter.action) || '';
    var result = routeAction(action, payload, 'POST');
    return buildResponse(result);
  } catch (err) {
    return buildResponse({ success: false, error: err.message });
  }
}

function buildResponse(data) {
  var output = ContentService.createTextOutput(JSON.stringify(data));
  output.setMimeType(ContentService.MimeType.JSON);
  return output;
}

// =========================================================
// Router & Permission Controller
// =========================================================

function routeAction(action, params, method) {
  params = params || {};
  var user = getCurrentUser(params);

  switch (action) {
    case 'ping':             return { success: true, message: 'pong', user: user };
    case 'getCurrentUser':   return { success: true, data: user };
    case 'getEvents':        return getEvents(params, user);
    case 'createEvent':      return createEvent(params, user);
    case 'updateEvent':      return updateEvent(params, user);
    case 'upsertEvent':      return upsertEvent(params, user); // สำหรับ Sync จาก Firestore
    case 'batchUpsertEvents':return batchUpsertEvents(params.events, user);
    case 'deleteEvent':      return deleteEvent(params.id, user);
    case 'updateStatus':     return updateStatus(params.id, params.status, user);
    case 'getUsers':         return getUsers(user);
    case 'createUser':       return createUser(params, user);
    case 'updateUser':       return updateUser(params, user);
    case 'deleteUser':       return deleteUser(params.email, user);
    case 'getHolidays':      return getHolidays(params.year);
    case 'createHoliday':    return createHoliday(params, user);
    case 'deleteHoliday':    return deleteHoliday(params.date, user);
    case 'getDashboard':     return getDashboard(user);
    case 'initSampleData':   return initSampleData();
    case 'setupTrigger':     return setupTrigger();
    default:                 return { success: false, error: 'Unknown action: ' + action };
  }
}

// =========================================================
// Authentication & User Helper
// =========================================================

function getCurrentUser(params) {
  try {
    var email = '';
    try {
      email = Session.getActiveUser().getEmail();
    } catch (sessionErr) {}

    // Fallback: รับ email ที่ส่งมาจาก frontend (Client-side authenticated)
    if (!email && params && params.userEmail) {
      email = String(params.userEmail).trim();
    }

    if (!email) {
      return { email: 'guest@system.local', name: 'ผู้เยี่ยมชม', department: '', role: 'staff' };
    }

    var sheet = getSheet(SHEET_USERS);
    ensureHeaders(sheet, USER_COLS);
    var users = sheetToObjects(sheet, USER_COLS);

    for (var i = 0; i < users.length; i++) {
      if (users[i].email.toLowerCase() === email.toLowerCase()) {
        return users[i];
      }
    }

    return {
      email: email,
      name: email.split('@')[0],
      department: '',
      role: 'staff'
    };
  } catch (e) {
    return { email: 'guest@system.local', name: 'ผู้เยี่ยมชม', department: '', role: 'staff' };
  }
}

function requireRole(user, allowedRoles) {
  if (!user || !user.role) {
    throw new Error('กรุณาเข้าสู่ระบบก่อนดำเนินการ');
  }
  if (allowedRoles.indexOf(user.role) === -1) {
    throw new Error('ไม่มีสิทธิ์ดำเนินการนี้ (ต้องการสิทธิ์: ' + allowedRoles.join('/') + ')');
  }
}

// =========================================================
// Spreadsheet Data Helpers
// =========================================================

function getSpreadsheet() {
  return SpreadsheetApp.getActiveSpreadsheet();
}

function getSheet(name) {
  var ss = getSpreadsheet();
  var sheet = ss.getSheetByName(name);
  if (!sheet) {
    sheet = ss.insertSheet(name);
  }
  return sheet;
}

function ensureHeaders(sheet, cols) {
  if (sheet.getLastRow() === 0) {
    sheet.appendRow(cols);
    return;
  }
  var firstRow = sheet.getRange(1, 1, 1, cols.length).getValues()[0];
  var needsHeader = firstRow.every(function(v) { return v === '' || v === null; });
  if (needsHeader) {
    sheet.getRange(1, 1, 1, cols.length).setValues([cols]);
  }
}

function sheetToObjects(sheet, cols) {
  var lastRow = sheet.getLastRow();
  var lastCol = cols.length;
  if (lastRow <= 1) return [];

  var data = sheet.getRange(1, 1, lastRow, lastCol).getValues();
  var result = [];

  for (var i = 1; i < data.length; i++) {
    var row = data[i];
    var obj = {};
    var isEmpty = true;

    for (var j = 0; j < cols.length; j++) {
      var val = row[j];
      if (val instanceof Date) {
        val = Utilities.formatDate(val, 'Asia/Bangkok', 'yyyy-MM-dd HH:mm:ss');
      }
      obj[cols[j]] = val !== undefined ? String(val).trim() : '';
      if (val !== '' && val !== null && val !== undefined) {
        isEmpty = false;
      }
    }
    if (!isEmpty) {
      result.push(obj);
    }
  }
  return result;
}

function objectToRow(obj, cols) {
  return cols.map(function(c) {
    return obj[c] !== undefined && obj[c] !== null ? obj[c] : '';
  });
}

function findRowIndex(sheet, colIndex, value) {
  var data = sheet.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][colIndex]).trim() === String(value).trim()) {
      return i + 1; // 1-based index
    }
  }
  return -1;
}

function generateUUID() {
  return Utilities.getUuid().replace(/-/g, '').substring(0, 16);
}

function getCurrentTimestamp() {
  return Utilities.formatDate(new Date(), 'Asia/Bangkok', 'yyyy-MM-dd HH:mm:ss');
}

// =========================================================
// Events Management & Sync
// =========================================================

function getEvents(params, user) {
  try {
    var sheet = getSheet(SHEET_EVENTS);
    ensureHeaders(sheet, EVENT_COLS);
    var events = sheetToObjects(sheet, EVENT_COLS);

    // กรองตาม Assignee
    if (params && params.assignee && params.assignee !== '') {
      events = events.filter(function(e) { return e.assignee === params.assignee; });
    }

    // กรองตาม Type
    if (params && params.type && params.type !== '') {
      events = events.filter(function(e) { return e.type === params.type; });
    }

    // กรองตาม Status
    if (params && params.status && params.status !== '') {
      events = events.filter(function(e) { return e.status === params.status; });
    }

    return { success: true, data: events };
  } catch (err) {
    return { success: false, error: err.message };
  }
}

// ฟังก์ชันหลักในการ Sync งานจาก Firestore ลงใน Google Sheets
function upsertEvent(data, user) {
  try {
    if (!data.id) throw new Error('ไม่พบ ID ของงาน');
    var sheet = getSheet(SHEET_EVENTS);
    ensureHeaders(sheet, EVENT_COLS);

    var eventObj = {
      id:        data.id,
      title:     data.title     || '',
      type:      data.type      || 'other',
      status:    data.status    || 'todo',
      assignee:  data.assignee  || '',
      start:     data.start     || '',
      end:       data.end       || data.start || '',
      startTime: data.startTime || '',
      endTime:   data.endTime   || '',
      location:  data.location  || '',
      detail:    data.detail    || '',
      updatedAt: data.updatedAt || getCurrentTimestamp()
    };

    var rowIndex = findRowIndex(sheet, 0, data.id);
    if (rowIndex !== -1) {
      // มีอยู่แล้ว ให้อัปเดตแถวเดิม
      sheet.getRange(rowIndex, 1, 1, EVENT_COLS.length).setValues([objectToRow(eventObj, EVENT_COLS)]);
      return { success: true, action: 'updated', data: eventObj };
    } else {
      // งานใหม่ ให้เพิ่มแถวใหม่
      sheet.appendRow(objectToRow(eventObj, EVENT_COLS));
      return { success: true, action: 'inserted', data: eventObj };
    }
  } catch (err) {
    return { success: false, error: err.message };
  }
}

function batchUpsertEvents(events, user) {
  try {
    if (!events || !Array.isArray(events)) throw new Error('ข้อมูล events ไม่ถูกต้อง');
    var sheet = getSheet(SHEET_EVENTS);
    ensureHeaders(sheet, EVENT_COLS);

    var count = 0;
    for (var i = 0; i < events.length; i++) {
      var res = upsertEvent(events[i], user);
      if (res.success) count++;
    }
    return { success: true, count: count, message: 'ซิงก์ข้อมูล ' + count + ' รายการสำเร็จ' };
  } catch (err) {
    return { success: false, error: err.message };
  }
}

function createEvent(data, user) {
  data.id = data.id || generateUUID();
  data.updatedAt = data.updatedAt || getCurrentTimestamp();
  return upsertEvent(data, user);
}

function updateEvent(data, user) {
  data.updatedAt = getCurrentTimestamp();
  return upsertEvent(data, user);
}

function deleteEvent(id, user) {
  try {
    if (!id) throw new Error('ไม่พบรหัสงาน (ID)');
    var sheet = getSheet(SHEET_EVENTS);
    ensureHeaders(sheet, EVENT_COLS);

    var rowIndex = findRowIndex(sheet, 0, id);
    if (rowIndex === -1) {
      return { success: true, message: 'ไม่พบงานในชีต (อาจถูกลบไปแล้ว)' };
    }

    sheet.deleteRow(rowIndex);
    return { success: true, message: 'ลบงานออกจากชีตเรียบร้อยแล้ว' };
  } catch (err) {
    return { success: false, error: err.message };
  }
}

function updateStatus(id, status, user) {
  try {
    if (!id) throw new Error('ไม่พบรหัสงาน (ID)');
    var validStatuses = ['todo', 'doing', 'done'];
    if (validStatuses.indexOf(status) === -1) {
      throw new Error('สถานะงานไม่ถูกต้อง');
    }

    var sheet = getSheet(SHEET_EVENTS);
    ensureHeaders(sheet, EVENT_COLS);

    var rowIndex = findRowIndex(sheet, 0, id);
    if (rowIndex === -1) throw new Error('ไม่พบแถวข้อมูลในชีต');

    var statusColIndex = EVENT_COLS.indexOf('status');
    var updatedColIndex = EVENT_COLS.indexOf('updatedAt');

    sheet.getRange(rowIndex, statusColIndex + 1).setValue(status);
    sheet.getRange(rowIndex, updatedColIndex + 1).setValue(getCurrentTimestamp());

    return { success: true, message: 'อัปเดตสถานะงานในชีตเรียบร้อยแล้ว' };
  } catch (err) {
    return { success: false, error: err.message };
  }
}

// =========================================================
// Users Management
// =========================================================

function getUsers(user) {
  try {
    var sheet = getSheet(SHEET_USERS);
    ensureHeaders(sheet, USER_COLS);
    var users = sheetToObjects(sheet, USER_COLS);
    return { success: true, data: users };
  } catch (err) {
    return { success: false, error: err.message };
  }
}

function createUser(data, user) {
  try {
    if (!data.email) throw new Error('กรุณาระบุอีเมล');
    if (!data.name)  throw new Error('กรุณาระบุชื่อ-นามสกุล');

    var sheet = getSheet(SHEET_USERS);
    ensureHeaders(sheet, USER_COLS);
    var existing = sheetToObjects(sheet, USER_COLS);

    for (var i = 0; i < existing.length; i++) {
      if (existing[i].email.toLowerCase() === String(data.email).trim().toLowerCase()) {
        throw new Error('อีเมลนี้มีอยู่ในระบบแล้ว');
      }
    }

    var newUser = {
      email:      String(data.email).trim(),
      name:       String(data.name).trim(),
      department: data.department ? String(data.department).trim() : '',
      role:       data.role || 'staff'
    };

    sheet.appendRow(objectToRow(newUser, USER_COLS));
    return { success: true, data: newUser, message: 'เพิ่มผู้ใช้สำเร็จ' };
  } catch (err) {
    return { success: false, error: err.message };
  }
}

function updateUser(data, user) {
  try {
    if (!data.email) throw new Error('กรุณาระบุอีเมล');

    var sheet = getSheet(SHEET_USERS);
    ensureHeaders(sheet, USER_COLS);

    var rowIndex = findRowIndex(sheet, 0, data.email);
    if (rowIndex === -1) throw new Error('ไม่พบผู้ใช้ที่ต้องการแก้ไข');

    var updatedUser = {
      email:      data.email,
      name:       data.name       || '',
      department: data.department || '',
      role:       data.role       || 'staff'
    };

    sheet.getRange(rowIndex, 1, 1, USER_COLS.length).setValues([objectToRow(updatedUser, USER_COLS)]);
    return { success: true, data: updatedUser, message: 'แก้ไขข้อมูลผู้ใช้สำเร็จ' };
  } catch (err) {
    return { success: false, error: err.message };
  }
}

function deleteUser(email, user) {
  try {
    if (!email) throw new Error('กรุณาระบุอีเมล');

    var sheet = getSheet(SHEET_USERS);
    ensureHeaders(sheet, USER_COLS);

    var rowIndex = findRowIndex(sheet, 0, email);
    if (rowIndex === -1) throw new Error('ไม่พบผู้ใช้ที่ต้องการลบ');

    sheet.deleteRow(rowIndex);
    return { success: true, message: 'ลบผู้ใช้เรียบร้อยแล้ว' };
  } catch (err) {
    return { success: false, error: err.message };
  }
}

// =========================================================
// Holidays Management
// =========================================================

function getHolidays(year) {
  try {
    var sheet = getSheet(SHEET_HOLIDAYS);
    ensureHeaders(sheet, HOLIDAY_COLS);
    var holidays = sheetToObjects(sheet, HOLIDAY_COLS);

    if (year) {
      var yearStr = String(year);
      holidays = holidays.filter(function(h) {
        return String(h.date).substring(0, 4) === yearStr;
      });
    }

    return { success: true, data: holidays };
  } catch (err) {
    return { success: false, error: err.message };
  }
}

function createHoliday(data, user) {
  try {
    if (!data.date) throw new Error('กรุณาระบุวันที่วันหยุด');
    if (!data.name) throw new Error('กรุณาระบุชื่อวันหยุด');

    var sheet = getSheet(SHEET_HOLIDAYS);
    ensureHeaders(sheet, HOLIDAY_COLS);
    var existing = sheetToObjects(sheet, HOLIDAY_COLS);

    for (var i = 0; i < existing.length; i++) {
      if (String(existing[i].date) === String(data.date)) {
        throw new Error('วันที่นี้มีรายการวันหยุดอยู่แล้ว');
      }
    }

    var holiday = {
      date: data.date,
      name: data.name,
      type: data.type || 'public'
    };

    sheet.appendRow(objectToRow(holiday, HOLIDAY_COLS));
    return { success: true, data: holiday, message: 'เพิ่มวันหยุดสำเร็จ' };
  } catch (err) {
    return { success: false, error: err.message };
  }
}

function deleteHoliday(date, user) {
  try {
    if (!date) throw new Error('กรุณาระบุวันที่');

    var sheet = getSheet(SHEET_HOLIDAYS);
    ensureHeaders(sheet, HOLIDAY_COLS);

    var rowIndex = findRowIndex(sheet, 0, date);
    if (rowIndex === -1) throw new Error('ไม่พบวันหยุดที่ต้องการลบ');

    sheet.deleteRow(rowIndex);
    return { success: true, message: 'ลบวันหยุดเรียบร้อยแล้ว' };
  } catch (err) {
    return { success: false, error: err.message };
  }
}

// =========================================================
// Dashboard & Analytics
// =========================================================

function getDashboard(user) {
  try {
    var eventSheet = getSheet(SHEET_EVENTS);
    var userSheet  = getSheet(SHEET_USERS);
    ensureHeaders(eventSheet, EVENT_COLS);
    ensureHeaders(userSheet, USER_COLS);

    var events = sheetToObjects(eventSheet, EVENT_COLS);
    var users  = sheetToObjects(userSheet, USER_COLS);
    var today  = Utilities.formatDate(new Date(), 'Asia/Bangkok', 'yyyy-MM-dd');

    var total   = events.length;
    var todo    = events.filter(function(e) { return e.status === 'todo'; }).length;
    var doing   = events.filter(function(e) { return e.status === 'doing'; }).length;
    var done    = events.filter(function(e) { return e.status === 'done'; }).length;
    var overdue = events.filter(function(e) {
      return e.status !== 'done' && e.end && e.end < today;
    }).length;

    var userSummary = users.map(function(u) {
      var userEvents = events.filter(function(e) { return e.assignee === u.email; });
      return {
        email:      u.email,
        name:       u.name,
        department: u.department,
        todo:       userEvents.filter(function(e) { return e.status === 'todo'; }).length,
        doing:      userEvents.filter(function(e) { return e.status === 'doing'; }).length,
        done:       userEvents.filter(function(e) { return e.status === 'done'; }).length,
        total:      userEvents.length
      };
    });

    var d7 = new Date();
    d7.setDate(d7.getDate() + 7);
    var d7Str = Utilities.formatDate(d7, 'Asia/Bangkok', 'yyyy-MM-dd');

    var upcoming = events.filter(function(e) {
      return e.status !== 'done' && e.end >= today && e.end <= d7Str;
    }).sort(function(a, b) {
      return a.end < b.end ? -1 : 1;
    });

    return {
      success: true,
      data: {
        summary: { total: total, todo: todo, doing: doing, done: done, overdue: overdue },
        userSummary: userSummary,
        upcoming: upcoming
      }
    };
  } catch (err) {
    return { success: false, error: err.message };
  }
}

// =========================================================
// Email Reminder (Time-Driven Trigger Function)
// =========================================================

function sendEmailReminder() {
  try {
    var eventSheet = getSheet(SHEET_EVENTS);
    var userSheet  = getSheet(SHEET_USERS);
    ensureHeaders(eventSheet, EVENT_COLS);
    ensureHeaders(userSheet, USER_COLS);

    var events = sheetToObjects(eventSheet, EVENT_COLS);
    var users  = sheetToObjects(userSheet, USER_COLS);

    var today = new Date();
    today.setHours(0, 0, 0, 0);
    var d3 = new Date(today);
    d3.setDate(d3.getDate() + 3);

    var todayStr = Utilities.formatDate(today, 'Asia/Bangkok', 'yyyy-MM-dd');
    var d3Str    = Utilities.formatDate(d3, 'Asia/Bangkok', 'yyyy-MM-dd');

    var dueEvents = events.filter(function(e) {
      return e.status !== 'done' && e.end >= todayStr && e.end <= d3Str;
    });

    if (dueEvents.length === 0) return;

    var grouped = {};
    dueEvents.forEach(function(e) {
      if (!grouped[e.assignee]) grouped[e.assignee] = [];
      grouped[e.assignee].push(e);
    });

    Object.keys(grouped).forEach(function(email) {
      if (!email || email.indexOf('@') === -1) return;

      var userEvents = grouped[email];
      var userName = email;
      for (var i = 0; i < users.length; i++) {
        if (users[i].email === email) {
          userName = users[i].name;
          break;
        }
      }

      var subject = '[ระบบปฏิทินการปฏิบัติงาน] แจ้งเตือนงานใกล้ถึงกำหนด (' + userEvents.length + ' รายการ)';
      var body = 'เรียน คุณ' + userName + ',\n\n' +
        'ระบบขอแจ้งเตือนงานที่ใกล้ถึงกำหนดส่ง/ปฏิบัติการภายใน 3 วัน ดังต่อไปนี้:\n\n';

      userEvents.forEach(function(ev, idx) {
        var statusThai = { todo: 'รอดำเนินการ', doing: 'กำลังดำเนินการ' }[ev.status] || ev.status;
        body += (idx + 1) + '. ' + ev.title + '\n' +
                '   - กำหนดส่ง/สิ้นสุด: ' + ev.end + (ev.endTime ? ' ' + ev.endTime + ' น.' : '') + '\n' +
                '   - สถานะ: ' + statusThai + '\n' +
                (ev.location ? '   - สถานที่: ' + ev.location + '\n' : '') +
                (ev.detail ? '   - รายละเอียด: ' + ev.detail + '\n' : '') + '\n';
      });

      body += 'กรุณาตรวจสอบและดำเนินการตามกำหนดเวลา\n\n' +
              'ขอแสดงความนับถือ,\n' +
              'ระบบปฏิทินการปฏิบัติงานอัตโนมัติ';

      try {
        MailApp.sendEmail(email, subject, body);
      } catch (mailErr) {
        Logger.log('ส่งอีเมลถึง ' + email + ' ล้มเหลว: ' + mailErr.message);
      }
    });

    return { success: true, count: dueEvents.length };
  } catch (err) {
    Logger.log('sendEmailReminder error: ' + err.message);
    return { success: false, error: err.message };
  }
}

function setupTrigger() {
  try {
    var triggers = ScriptApp.getProjectTriggers();
    triggers.forEach(function(t) {
      if (t.getHandlerFunction() === 'sendEmailReminder') {
        ScriptApp.deleteTrigger(t);
      }
    });

    ScriptApp.newTrigger('sendEmailReminder')
      .timeBased()
      .everyDays(1)
      .atHour(8)
      .create();

    return {
      success: true,
      message: 'ตั้งค่า Daily Trigger สำเร็จ (ทำงานทุกวัน เวลา 08:00 น.)'
    };
  } catch (err) {
    return { success: false, error: err.message };
  }
}

// =========================================================
// Sample Data Initialization
// =========================================================

function initSampleData() {
  try {
    // 1. Users
    var userSheet = getSheet(SHEET_USERS);
    userSheet.clearContents();
    userSheet.appendRow(USER_COLS);

    var sampleUsers = [
      ['admin@school.ac.th',   'ผู้ดูแลระบบ',    'IT', 'admin'],
      ['head.it@school.ac.th', 'หัวหน้าฝ่าย IT', 'IT', 'head'],
      ['head.hr@school.ac.th', 'หัวหน้าฝ่าย HR', 'HR', 'head'],
      ['staff1@school.ac.th',  'นายสมชาย ใจดี',  'IT', 'staff'],
      ['staff2@school.ac.th',  'นางสมหญิง รักดี', 'HR', 'staff']
    ];
    sampleUsers.forEach(function(row) { userSheet.appendRow(row); });

    // 2. Holidays
    var holSheet = getSheet(SHEET_HOLIDAYS);
    holSheet.clearContents();
    holSheet.appendRow(HOLIDAY_COLS);

    var sampleHolidays = [
      ['2025-01-01', 'วันขึ้นปีใหม่',                           'public'],
      ['2025-02-12', 'วันมาฆบูชา',                               'public'],
      ['2025-04-06', 'วันจักรี',                                 'public'],
      ['2025-04-13', 'วันสงกรานต์',                              'public'],
      ['2025-04-14', 'วันสงกรานต์',                              'public'],
      ['2025-04-15', 'วันสงกรานต์',                              'public'],
      ['2025-05-01', 'วันแรงงานแห่งชาติ',                        'public'],
      ['2025-05-05', 'วันฉัตรมงคล',                              'public'],
      ['2025-05-12', 'วันวิสาขบูชา',                             'public'],
      ['2025-06-03', 'วันเฉลิมพระชนมพรรษาสมเด็จพระนางเจ้าฯ พระบรมราชินี', 'public'],
      ['2025-07-10', 'วันอาสาฬหบูชา',                           'public'],
      ['2025-07-11', 'วันเข้าพรรษา',                             'public'],
      ['2025-07-28', 'วันเฉลิมพระชนมพรรษาพระบาทสมเด็จพระเจ้าอยู่หัว', 'public'],
      ['2025-08-12', 'วันแม่แห่งชาติ',                           'public'],
      ['2025-10-13', 'วันนวมินทรมหาราช',                         'public'],
      ['2025-10-23', 'วันปิยมหาราช',                             'public'],
      ['2025-12-05', 'วันคล้ายวันพระบรมราชสมภพ ร.9 / วันพ่อแห่งชาติ', 'public'],
      ['2025-12-10', 'วันรัฐธรรมนูญ',                            'public'],
      ['2025-12-31', 'วันสิ้นปี',                                'public']
    ];
    sampleHolidays.forEach(function(row) { holSheet.appendRow(row); });

    // 3. Events (พร้อม updatedAt)
    var evSheet = getSheet(SHEET_EVENTS);
    evSheet.clearContents();
    evSheet.appendRow(EVENT_COLS);

    var nowStr = getCurrentTimestamp();
    var sampleEvents = [
      [generateUUID(), 'ประชุมวางแผนโครงการประจำปี 2568',         'meet',  'done',  'head.it@school.ac.th', '2025-01-10', '2025-01-10', '09:00', '12:00', 'ห้องประชุม 1',     'ประชุมวางแผนงานประจำปีงบประมาณ 2568', nowStr],
      [generateUUID(), 'อบรมเชิงปฏิบัติการระบบ ERP องค์กร',        'teach', 'done',  'staff1@school.ac.th',  '2025-01-20', '2025-01-22', '08:30', '16:30', 'ห้องคอมพิวเตอร์ 3', 'อบรมหลักสูตรการใช้งานระบบ ERP สำหรับบุคลากรใหม่', nowStr],
      [generateUUID(), 'โครงการพัฒนาเว็บไซต์และพอร์ทัลสถาบัน',     'proj',  'doing', 'staff1@school.ac.th',  '2025-02-01', '2025-03-31', '08:30', '16:30', 'ฝ่ายสารสนเทศ IT',  'ปรับปรุงโฉมและฟังก์ชันเว็บไซต์หลักของสถาบัน', nowStr],
      [generateUUID(), 'ประชุมคณะกรรมการบริหาร ประจำไตรมาส 1',    'meet',  'done',  'head.hr@school.ac.th', '2025-02-15', '2025-02-15', '13:00', '16:00', 'ห้องประชุมใหญ่',   'รายงานสรุปผลการดำเนินงานและเป้าหมายไตรมาสถัดไป', nowStr],
      [generateUUID(), 'อบรม Microsoft 365 เพื่อการทำงานร่วมกัน', 'teach', 'todo',  'staff2@school.ac.th',  '2025-03-10', '2025-03-12', '08:30', '12:00', 'ห้องอบรม 2',       'การใช้งาน Teams, OneDrive และ Forms เพื่อการทำงาน', nowStr],
      [generateUUID(), 'โครงการสำรวจความผูกพันและความพึงพอใจ',     'proj',  'todo',  'staff2@school.ac.th',  '2025-03-15', '2025-04-30', '',      '',      'ฝ่ายทรัพยากรบุคคล', 'จัดทำแบบสำรวจและประมวลผลข้อมูลความพึงพอใจบุคลากร', nowStr],
      [generateUUID(), 'ประชุมปฐมนิเทศผู้ปกครอง ภาคเรียนที่ 1/2568', 'meet',  'todo',  'head.hr@school.ac.th', '2025-05-20', '2025-05-20', '14:00', '17:00', 'หอประชุมใหญ่',     'ชี้แจงแนวทางการเรียนการสอนและกฎระเบียบของสถาบัน', nowStr],
      [generateUUID(), 'อบรมการตระหนักรู้ด้านความปลอดภัยไซเบอร์', 'teach', 'doing', 'head.it@school.ac.th', '2025-06-05', '2025-06-06', '09:00', '16:00', 'ห้องประชุม 2',     'หลักสูตร Cybersecurity Awareness สำหรับเจ้าหน้าที่ทุกฝ่าย', nowStr],
      [generateUUID(), 'โครงการจัดทำคลังข้อมูลและแฟ้มประวัติบุคคล', 'proj',  'doing', 'staff1@school.ac.th',  '2025-07-01', '2025-08-31', '08:30', '16:30', 'ฝ่ายบุคคล HR',     'แปลงเอกสารสู่ระบบดิจิทัลและจัดหมวดหมู่ข้อมูล', nowStr],
      [generateUUID(), 'กิจกรรมเฉลิมพระเกียรติวันแม่แห่งชาติ',     'other', 'todo',  'staff2@school.ac.th',  '2025-08-12', '2025-08-12', '08:00', '12:00', 'ลานอเนกประสงค์',   'พิธีถวายพระพรชัยมงคลและกิจกรรมบำเพ็ญประโยชน์', nowStr]
    ];
    sampleEvents.forEach(function(row) { evSheet.appendRow(row); });

    return {
      success: true,
      message: 'สร้างและตั้งค่าข้อมูลตัวอย่างเรียบร้อยแล้ว (5 ผู้ใช้, 19 วันหยุด, 10 งาน พร้อม updatedAt)'
    };
  } catch (err) {
    return { success: false, error: err.message };
  }
}
