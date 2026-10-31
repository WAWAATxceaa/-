// =========================================================
// ระบบปฏิทินการปฏิบัติงาน - Google Apps Script Backend
// Code.gs
// =========================================================

// ชื่อ Sheet ใน Google Spreadsheet
var SHEET_EVENTS   = 'Events';
var SHEET_USERS    = 'Users';
var SHEET_HOLIDAYS = 'Holidays';

// กำหนดคอลัมน์ตามข้อกำหนด
var EVENT_COLS   = ['id','title','type','status','assignee','start','end','startTime','endTime','location','detail'];
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
      // หากรันเป็น API เพียวๆ โดยไม่มีไฟล์ index.html ใน GAS
      return buildResponse({
        success: true,
        message: 'ระบบปฏิทินการปฏิบัติงาน API กำลังทำงาน',
        version: '1.0'
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
    case 'ping':           return { success: true, message: 'pong', user: user };
    case 'getCurrentUser': return { success: true, data: user };
    case 'getEvents':      return getEvents(params, user);
    case 'createEvent':    return createEvent(params, user);
    case 'updateEvent':    return updateEvent(params, user);
    case 'deleteEvent':    return deleteEvent(params.id, user);
    case 'updateStatus':   return updateStatus(params.id, params.status, user);
    case 'getUsers':       return getUsers(user);
    case 'createUser':     return createUser(params, user);
    case 'updateUser':     return updateUser(params, user);
    case 'deleteUser':     return deleteUser(params.email, user);
    case 'getHolidays':    return getHolidays(params.year);
    case 'createHoliday':  return createHoliday(params, user);
    case 'deleteHoliday':  return deleteHoliday(params.date, user);
    case 'getDashboard':   return getDashboard(user);
    case 'initSampleData': return initSampleData();
    case 'setupTrigger':   return setupTrigger();
    default:               return { success: false, error: 'Unknown action: ' + action };
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

    // หากไม่พบในฐานข้อมูลชีต กำหนดค่าเริ่มต้นเป็น staff
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
        val = Utilities.formatDate(val, 'Asia/Bangkok', 'yyyy-MM-dd');
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

// =========================================================
// Events Management
// =========================================================

function getEvents(params, user) {
  try {
    var sheet = getSheet(SHEET_EVENTS);
    ensureHeaders(sheet, EVENT_COLS);
    var events = sheetToObjects(sheet, EVENT_COLS);

    // กรองตามสิทธิ์:
    // head เห็นทั้งแผนกของตน
    if (user && user.role === 'head' && user.department) {
      var usersSheet = getSheet(SHEET_USERS);
      ensureHeaders(usersSheet, USER_COLS);
      var allUsers = sheetToObjects(usersSheet, USER_COLS);
      var deptEmails = allUsers
        .filter(function(u) { return u.department === user.department; })
        .map(function(u) { return u.email; });

      // ถ้าไม่ได้ระบุฟิลเตอร์เจาะจง ให้เห็นงานในแผนกและงานส่วนกลาง
      if (params.departmentOnly === 'true') {
        events = events.filter(function(e) {
          return deptEmails.indexOf(e.assignee) !== -1 || e.assignee === user.email;
        });
      }
    }

    // กรองตาม Assignee
    if (params.assignee && params.assignee !== '') {
      events = events.filter(function(e) { return e.assignee === params.assignee; });
    }

    // กรองตาม Type (teach, meet, proj, hol, other)
    if (params.type && params.type !== '') {
      events = events.filter(function(e) { return e.type === params.type; });
    }

    // กรองตาม Status (todo, doing, done)
    if (params.status && params.status !== '') {
      events = events.filter(function(e) { return e.status === params.status; });
    }

    // กรองตามช่วงวันที่
    if (params.startDate && params.endDate) {
      events = events.filter(function(e) {
        return (e.start <= params.endDate && (e.end || e.start) >= params.startDate);
      });
    }

    return { success: true, data: events };
  } catch (err) {
    return { success: false, error: err.message };
  }
}

function createEvent(data, user) {
  try {
    if (!data.title) throw new Error('กรุณาระบุชื่องาน');
    if (!data.start) throw new Error('กรุณาระบุวันเริ่มต้น');

    var sheet = getSheet(SHEET_EVENTS);
    ensureHeaders(sheet, EVENT_COLS);

    var event = {
      id:        generateUUID(),
      title:     data.title     || '',
      type:      data.type      || 'other',
      status:    data.status    || 'todo',
      assignee:  data.assignee  || (user ? user.email : ''),
      start:     data.start     || '',
      end:       data.end       || data.start || '',
      startTime: data.startTime || '',
      endTime:   data.endTime   || '',
      location:  data.location  || '',
      detail:    data.detail    || ''
    };

    sheet.appendRow(objectToRow(event, EVENT_COLS));
    return { success: true, data: event, message: 'บันทึกงานใหม่สำเร็จ' };
  } catch (err) {
    return { success: false, error: err.message };
  }
}

function updateEvent(data, user) {
  try {
    if (!data.id) throw new Error('ไม่พบรหัสงาน (ID)');
    var sheet = getSheet(SHEET_EVENTS);
    ensureHeaders(sheet, EVENT_COLS);
    var allEvents = sheetToObjects(sheet, EVENT_COLS);

    var existingEvent = null;
    for (var i = 0; i < allEvents.length; i++) {
      if (allEvents[i].id === data.id) {
        existingEvent = allEvents[i];
        break;
      }
    }
    if (!existingEvent) throw new Error('ไม่พบข้อมูลงานในระบบ');

    // ตรวจสอบสิทธิ์:
    // admin: แก้ไขได้ทั้งหมด
    // head: แก้ไขงานในแผนกตนเองได้
    // staff: แก้ไขเฉพาะงานของตนเองเท่านั้น
    if (user.role === 'staff' && existingEvent.assignee !== user.email) {
      throw new Error('สิทธิ์เจ้าหน้าที่ (staff) สามารถแก้ไขเฉพาะงานของตนเองเท่านั้น');
    }
    if (user.role === 'head' && user.department) {
      var usersSheet = getSheet(SHEET_USERS);
      var allUsers = sheetToObjects(usersSheet, USER_COLS);
      var deptEmails = allUsers
        .filter(function(u) { return u.department === user.department; })
        .map(function(u) { return u.email; });
      if (deptEmails.indexOf(existingEvent.assignee) === -1 && existingEvent.assignee !== user.email && user.role !== 'admin') {
        throw new Error('หัวหน้าแผนกสามารถแก้ไขเฉพาะงานในแผนกตนเองเท่านั้น');
      }
    }

    var updatedEvent = {
      id:        data.id,
      title:     data.title     !== undefined ? data.title     : existingEvent.title,
      type:      data.type      !== undefined ? data.type      : existingEvent.type,
      status:    data.status    !== undefined ? data.status    : existingEvent.status,
      assignee:  data.assignee  !== undefined ? data.assignee  : existingEvent.assignee,
      start:     data.start     !== undefined ? data.start     : existingEvent.start,
      end:       data.end       !== undefined ? data.end       : existingEvent.end,
      startTime: data.startTime !== undefined ? data.startTime : existingEvent.startTime,
      endTime:   data.endTime   !== undefined ? data.endTime   : existingEvent.endTime,
      location:  data.location  !== undefined ? data.location  : existingEvent.location,
      detail:    data.detail    !== undefined ? data.detail    : existingEvent.detail
    };

    var rowIndex = findRowIndex(sheet, 0, data.id);
    if (rowIndex === -1) throw new Error('ไม่พบแถวข้อมูลในชีต');

    sheet.getRange(rowIndex, 1, 1, EVENT_COLS.length).setValues([objectToRow(updatedEvent, EVENT_COLS)]);
    return { success: true, data: updatedEvent, message: 'แก้ไขข้อมูลงานสำเร็จ' };
  } catch (err) {
    return { success: false, error: err.message };
  }
}

function deleteEvent(id, user) {
  try {
    if (!id) throw new Error('ไม่พบรหัสงาน (ID)');
    var sheet = getSheet(SHEET_EVENTS);
    ensureHeaders(sheet, EVENT_COLS);
    var allEvents = sheetToObjects(sheet, EVENT_COLS);

    var targetEvent = null;
    for (var i = 0; i < allEvents.length; i++) {
      if (allEvents[i].id === id) {
        targetEvent = allEvents[i];
        break;
      }
    }
    if (!targetEvent) throw new Error('ไม่พบข้อมูลงานที่ต้องการลบ');

    // ตรวจสอบสิทธิ์
    if (user.role === 'staff' && targetEvent.assignee !== user.email) {
      throw new Error('สิทธิ์เจ้าหน้าที่ (staff) สามารถลบเฉพาะงานของตนเองเท่านั้น');
    }

    var rowIndex = findRowIndex(sheet, 0, id);
    if (rowIndex === -1) throw new Error('ไม่พบแถวข้อมูลในชีต');

    sheet.deleteRow(rowIndex);
    return { success: true, message: 'ลบงานเรียบร้อยแล้ว' };
  } catch (err) {
    return { success: false, error: err.message };
  }
}

function updateStatus(id, status, user) {
  try {
    if (!id) throw new Error('ไม่พบรหัสงาน (ID)');
    var validStatuses = ['todo', 'doing', 'done'];
    if (validStatuses.indexOf(status) === -1) {
      throw new Error('สถานะงานไม่ถูกต้อง (ต้องเป็น todo, doing หรือ done)');
    }

    var sheet = getSheet(SHEET_EVENTS);
    ensureHeaders(sheet, EVENT_COLS);
    var allEvents = sheetToObjects(sheet, EVENT_COLS);

    var targetEvent = null;
    for (var i = 0; i < allEvents.length; i++) {
      if (allEvents[i].id === id) {
        targetEvent = allEvents[i];
        break;
      }
    }
    if (!targetEvent) throw new Error('ไม่พบข้อมูลงาน');

    // ตรวจสอบสิทธิ์
    if (user.role === 'staff' && targetEvent.assignee !== user.email) {
      throw new Error('สิทธิ์เจ้าหน้าที่ (staff) สามารถเปลี่ยนสถานะเฉพาะงานของตนเองเท่านั้น');
    }

    var statusColIndex = EVENT_COLS.indexOf('status');
    var rowIndex = findRowIndex(sheet, 0, id);
    if (rowIndex === -1) throw new Error('ไม่พบแถวข้อมูลในชีต');

    sheet.getRange(rowIndex, statusColIndex + 1).setValue(status);
    targetEvent.status = status;

    return { success: true, data: targetEvent, message: 'อัปเดตสถานะงานเรียบร้อยแล้ว' };
  } catch (err) {
    return { success: false, error: err.message };
  }
}

// =========================================================
// Users Management (Admin)
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
    requireRole(user, ['admin']);
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
    requireRole(user, ['admin']);
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
    requireRole(user, ['admin']);
    if (!email) throw new Error('กรุณาระบุอีเมล');
    if (user.email === email) throw new Error('ไม่สามารถลบบัญชีของตนเองได้');

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
// Holidays Management (Admin)
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
    requireRole(user, ['admin']);
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
    requireRole(user, ['admin']);
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

    // Summary counts
    var total   = events.length;
    var todo    = events.filter(function(e) { return e.status === 'todo'; }).length;
    var doing   = events.filter(function(e) { return e.status === 'doing'; }).length;
    var done    = events.filter(function(e) { return e.status === 'done'; }).length;
    var overdue = events.filter(function(e) {
      return e.status !== 'done' && e.end && e.end < today;
    }).length;

    // Per-user summary
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

    // Upcoming events (ภายใน 7 วันข้างหน้า)
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

    // ค้นหางานที่ยังไม่เสร็จและจะครบกำหนดภายใน 3 วันข้างหน้า
    var dueEvents = events.filter(function(e) {
      return e.status !== 'done' && e.end >= todayStr && e.end <= d3Str;
    });

    if (dueEvents.length === 0) {
      Logger.log('ไม่มีงานที่ใกล้ถึงกำหนดภายใน 3 วัน');
      return;
    }

    // จัดกลุ่มตามอีเมลผู้รับผิดชอบ
    var grouped = {};
    dueEvents.forEach(function(e) {
      if (!grouped[e.assignee]) grouped[e.assignee] = [];
      grouped[e.assignee].push(e);
    });

    // ส่งอีเมลแจ้งเตือนรายบุคคล
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
        Logger.log('ส่งอีเมลแจ้งเตือนถึง ' + email + ' สำเร็จ');
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

// =========================================================
// Daily Trigger Setup
// =========================================================

function setupTrigger() {
  try {
    var triggers = ScriptApp.getProjectTriggers();
    triggers.forEach(function(t) {
      if (t.getHandlerFunction() === 'sendEmailReminder') {
        ScriptApp.deleteTrigger(t);
      }
    });

    // สร้าง Daily Trigger รันเวลา 08:00 น. ทุกวัน
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

    // 2. Holidays (วันหยุดนักขัตฤกษ์ไทย ปี พ.ศ. 2568 / ค.ศ. 2025)
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

    // 3. Events (งานตัวอย่าง ปี พ.ศ. 2568 / 2025)
    var evSheet = getSheet(SHEET_EVENTS);
    evSheet.clearContents();
    evSheet.appendRow(EVENT_COLS);

    var sampleEvents = [
      [generateUUID(), 'ประชุมวางแผนโครงการประจำปี 2568',         'meet',  'done',  'head.it@school.ac.th', '2025-01-10', '2025-01-10', '09:00', '12:00', 'ห้องประชุม 1',     'ประชุมวางแผนงานประจำปีงบประมาณ 2568 ร่วมกับคณะทำงาน'],
      [generateUUID(), 'อบรมเชิงปฏิบัติการระบบ ERP องค์กร',        'teach', 'done',  'staff1@school.ac.th',  '2025-01-20', '2025-01-22', '08:30', '16:30', 'ห้องคอมพิวเตอร์ 3', 'อบรมหลักสูตรการใช้งานระบบ ERP สำหรับบุคลากรใหม่'],
      [generateUUID(), 'โครงการพัฒนาเว็บไซต์และพอร์ทัลสถาบัน',     'proj',  'doing', 'staff1@school.ac.th',  '2025-02-01', '2025-03-31', '08:30', '16:30', 'ฝ่ายสารสนเทศ IT',  'ปรับปรุงโฉมและฟังก์ชันเว็บไซต์หลักของสถาบัน'],
      [generateUUID(), 'ประชุมคณะกรรมการบริหาร ประจำไตรมาส 1',    'meet',  'done',  'head.hr@school.ac.th', '2025-02-15', '2025-02-15', '13:00', '16:00', 'ห้องประชุมใหญ่',   'รายงานสรุปผลการดำเนินงานและเป้าหมายไตรมาสถัดไป'],
      [generateUUID(), 'อบรม Microsoft 365 เพื่อการทำงานร่วมกัน', 'teach', 'todo',  'staff2@school.ac.th',  '2025-03-10', '2025-03-12', '08:30', '12:00', 'ห้องอบรม 2',       'การใช้งาน Teams, OneDrive และ Forms เพื่อการทำงาน'],
      [generateUUID(), 'โครงการสำรวจความผูกพันและความพึงพอใจ',     'proj',  'todo',  'staff2@school.ac.th',  '2025-03-15', '2025-04-30', '',      '',      'ฝ่ายทรัพยากรบุคคล', 'จัดทำแบบสำรวจและประมวลผลข้อมูลความพึงพอใจบุคลากร'],
      [generateUUID(), 'ประชุมปฐมนิเทศผู้ปกครอง ภาคเรียนที่ 1/2568', 'meet',  'todo',  'head.hr@school.ac.th', '2025-05-20', '2025-05-20', '14:00', '17:00', 'หอประชุมใหญ่',     'ชี้แจงแนวทางการเรียนการสอนและกฎระเบียบของสถาบัน'],
      [generateUUID(), 'อบรมการตระหนักรู้ด้านความปลอดภัยไซเบอร์', 'teach', 'doing', 'head.it@school.ac.th', '2025-06-05', '2025-06-06', '09:00', '16:00', 'ห้องประชุม 2',     'หลักสูตร Cybersecurity Awareness สำหรับเจ้าหน้าที่ทุกฝ่าย'],
      [generateUUID(), 'โครงการจัดทำคลังข้อมูลและแฟ้มประวัติบุคคล', 'proj',  'doing', 'staff1@school.ac.th',  '2025-07-01', '2025-08-31', '08:30', '16:30', 'ฝ่ายบุคคล HR',     'แปลงเอกสารสู่ระบบดิจิทัลและจัดหมวดหมู่ข้อมูล'],
      [generateUUID(), 'กิจกรรมเฉลิมพระเกียรติวันแม่แห่งชาติ',     'other', 'todo',  'staff2@school.ac.th',  '2025-08-12', '2025-08-12', '08:00', '12:00', 'ลานอเนกประสงค์',   'พิธีถวายพระพรชัยมงคลและกิจกรรมบำเพ็ญประโยชน์']
    ];
    sampleEvents.forEach(function(row) { evSheet.appendRow(row); });

    return {
      success: true,
      message: 'สร้างและตั้งค่าข้อมูลตัวอย่างเรียบร้อยแล้ว (5 ผู้ใช้, 19 วันหยุด, 10 งาน)'
    };
  } catch (err) {
    return { success: false, error: err.message };
  }
}
