/**
 * ============================================================
 *  ระบบจัดตารางเวร "ห้องยาสบปราบ"
 *  Google Apps Script Backend — Complete Version
 * ============================================================
 *  วิธีใช้งาน:
 *  1. สร้าง Google Spreadsheet ใหม่
 *  2. เปิด Tools > Apps Script
 *  3. วางโค้ดทั้งหมดในไฟล์ Code.gs (แทนที่โค้ดเดิม)
 *  4. กด Run > initializeSheets หนึ่งครั้ง
 *  5. ตั้งค่า LINE Token ในชีท settings
 *  6. Deploy > New deployment > Web app > Anyone
 * ============================================================
 */

// ============ GLOBAL CONFIG ============
var BUDGET_OFFSET = 543; // พ.ศ. - ค.ศ.

// ============ SHEET HELPERS ============

function getSheet(name) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(name);
  if (!sheet) {
    sheet = ss.insertSheet(name);
  }
  return sheet;
}

/**
 * อ่านข้อมูลทั้งหมดจากชีท แปลงเป็น Array of Objects
 */
function getAllData(sheetName) {
  try {
    var sheet = getSheet(sheetName);
    var lastRow = sheet.getLastRow();
    var lastCol = sheet.getLastColumn();
    if (lastRow <= 1 || lastCol === 0) return [];
    var data = sheet.getRange(1, 1, lastRow, lastCol).getValues();
    if (data.length <= 1) return [];
    var headers = data[0];
    var result = [];
    for (var i = 1; i < data.length; i++) {
      var obj = {};
      for (var j = 0; j < headers.length; j++) {
        var val = data[i][j];
        // Convert Date objects to YYYY-MM-DD strings
        if (val instanceof Date && !isNaN(val)) {
          var y = val.getFullYear();
          var m = ('0' + (val.getMonth() + 1)).slice(-2);
          var d = ('0' + val.getDate()).slice(-2);
          val = y + '-' + m + '-' + d;
        }
        obj[String(headers[j])] = val;
      }
      result.push(obj);
    }
    return result;
  } catch(e) {
    Logger.log('getAllData error [' + sheetName + ']: ' + e.message);
    return [];
  }
}

/**
 * เพิ่มแถวใหม่ในชีท
 */
function addRow(sheetName, rowData) {
  var sheet = getSheet(sheetName);
  var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  var row = [];
  for (var i = 0; i < headers.length; i++) {
    var val = rowData[String(headers[i])];
    row.push(val === undefined || val === null ? '' : val);
  }
  sheet.appendRow(row);
  return true;
}

/**
 * อัพเดทแถวตามเงื่อนไข
 */
function updateRow(sheetName, matchField, matchValue, updateData) {
  var sheet = getSheet(sheetName);
  var data = sheet.getDataRange().getValues();
  if (data.length <= 1) return false;
  var headers = data[0];
  var matchIdx = headers.indexOf(matchField);
  if (matchIdx === -1) return false;

  var updateKeys = Object.keys(updateData);
  var updateCols = updateKeys.map(function(k) { return headers.indexOf(k); });

  for (var i = 1; i < data.length; i++) {
    if (String(data[i][matchIdx]).trim() === String(matchValue).trim()) {
      for (var k = 0; k < updateKeys.length; k++) {
        var colIdx = updateCols[k];
        if (colIdx >= 0) {
          sheet.getRange(i + 1, colIdx + 1).setValue(updateData[updateKeys[k]]);
        }
      }
      return true;
    }
  }
  return false;
}

/**
 * ลบแถวตามเงื่อนไข
 */
function deleteRow(sheetName, matchField, matchValue) {
  var sheet = getSheet(sheetName);
  var data = sheet.getDataRange().getValues();
  if (data.length <= 1) return false;
  var headers = data[0];
  var matchIdx = headers.indexOf(matchField);
  if (matchIdx === -1) return false;

  for (var i = data.length - 1; i >= 1; i--) {
    if (String(data[i][matchIdx]).trim() === String(matchValue).trim()) {
      sheet.deleteRow(i + 1);
      return true;
    }
  }
  return false;
}

/**
 * สร้าง ID แบบ Unique
 */
function generateId(prefix) {
  return prefix + '_' + Utilities.getUuid().substring(0, 8) + '_' + Date.now();
}

/**
 * แปลง Date object เป็น YYYY-MM-DD
 */
function formatDateKey(date) {
  if (!(date instanceof Date) || isNaN(date)) return '';
  var y = date.getFullYear();
  var m = ('0' + (date.getMonth() + 1)).slice(-2);
  var d = ('0' + date.getDate()).slice(-2);
  return y + '-' + m + '-' + d;
}

/**
 * แปลง Date object เป็นรูปแบบไทย วัน/เดือน/ปี พ.ศ.
 */
function formatDateThai(date) {
  if (!(date instanceof Date) || isNaN(date)) return '';
  var thaiMonths = ['ม.ค.','ก.พ.','มี.ค.','เม.ย.','พ.ค.','มิ.ย.','ก.ค.','ส.ค.','ก.ย.','ต.ค.','พ.ย.','ธ.ค.'];
  return date.getDate() + ' ' + thaiMonths[date.getMonth()] + ' พ.ศ. ' + (date.getFullYear() + BUDGET_OFFSET);
}

/**
 * แปลง Date object เป็นรูปแบบ ว/ด/พ.ศ.
 */
function formatDateNumeric(date) {
  if (!(date instanceof Date) || isNaN(date)) return '';
  var d = ('0' + date.getDate()).slice(-2);
  var m = ('0' + (date.getMonth() + 1)).slice(-2);
  var y = date.getFullYear() + BUDGET_OFFSET;
  return d + '/' + m + '/' + y;
}

/**
 * แปลง YYYY-MM-DD เป็น Date object
 */
function parseDateString(dateStr) {
  if (dateStr instanceof Date) return dateStr;
  if (!dateStr) return null;
  var str = String(dateStr);
  var parts = str.substring(0, 10).split('-');
  if (parts.length !== 3) return null;
  return new Date(parseInt(parts[0]), parseInt(parts[1]) - 1, parseInt(parts[2]));
}

/**
 * ชื่อเดือนไทย
 */
function getThaiMonthName(monthNum) {
  var months = ['','มกราคม','กุมภาพันธ์','มีนาคม','เมษายน','พฤษภาคม','มิถุนายน',
    'กรกฎาคม','สิงหาคม','กันยายน','ตุลาคม','พฤศจิกายน','ธันวาคม'];
  return months[parseInt(monthNum)] || '';
}

/**
 * ชื่อเดือนไทยแบบย่อ
 */
function getThaiMonthNameShort(monthIdx) {
  var names = ['ม.ค.','ก.พ.','มี.ค.','เม.ย.','พ.ค.','มิ.ย.','ก.ค.','ส.ค.','ก.ย.','ต.ค.','พ.ย.','ธ.ค.'];
  return names[parseInt(monthIdx)] || '';
}

// ============ INITIALIZE SHEETS ============

function initializeSheets() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();

  // --- 1. users ---
  var s = ss.getSheetByName('users');
  if (!s) {
    s = ss.insertSheet('users');
    s.getRange(1, 1, 1, 7).setValues([['user_id','fullname','position','rate_per_hour','line_user_id','is_admin','password']]);
    s.appendRow(['USR_001','Admin','ผู้ดูแลระบบ',0,'',true,'admin1234']);
  }

  // --- 2. duty_shifts ---
  s = ss.getSheetByName('duty_shifts');
  if (!s) {
    s = ss.insertSheet('duty_shifts');
    s.getRange(1, 1, 1, 8).setValues([['shift_id','shift_name','shift_type','hours','time_start','time_end','rate_multiplier','notify_enabled']]);
    s.appendRow(['SHIFT_NORMAL','เวรวันทำงานปกติ','วันปกติ',4,'16:00','20:00',1,true]);
    s.appendRow(['SHIFT_WEEKEND','เวรเสาร์-อาทิตย์/นักขัตฤกษ์','วันหยุด',8,'08:00','16:00',1.5,true]);
    s.appendRow(['SHIFT_DOC_NORMAL','เวรเอกสาร (วันปกติ)','เวรเอกสาร',4,'16:00','20:00',1,true]);
    s.appendRow(['SHIFT_DOC_WEEKEND','เวรเอกสาร (วันหยุด)','เวรเอกสาร',8,'08:30','16:30',1.5,true]);
  }

  // --- 3. bookings ---
  s = ss.getSheetByName('bookings');
  if (!s) {
    s = ss.insertSheet('bookings');
    s.getRange(1, 1, 1, 8).setValues([['booking_id','duty_date','shift_type','user_id','booked_by','booking_time','status','notes']]);
  }

  // --- 4. notifications ---
  s = ss.getSheetByName('notifications');
  if (!s) {
    s = ss.insertSheet('notifications');
    s.getRange(1, 1, 1, 4).setValues([['notif_id','user_id','shift_type','is_active']]);
  }

  // --- 5. holidays ---
  s = ss.getSheetByName('holidays');
  if (!s) {
    s = ss.insertSheet('holidays');
    s.getRange(1, 1, 1, 4).setValues([['holiday_date','holiday_name','holiday_type','year']]);
    // วันหยุดนักขัตฤกษ์ พ.ศ. 2568
    var hol = [
      ['2025-01-01','วันขึ้นปีใหม่','national_holiday',2568],
      ['2025-02-12','วันมาฆบูชา','national_holiday',2568],
      ['2025-04-06','วันจักรี','national_holiday',2568],
      ['2025-04-13','วันสงกรานต์','national_holiday',2568],
      ['2025-04-14','วันสงกรานต์','national_holiday',2568],
      ['2025-04-15','วันสงกรานต์','national_holiday',2568],
      ['2025-05-01','วันแรงงาน','national_holiday',2568],
      ['2025-05-05','วันฉัตรมงคล','national_holiday',2568],
      ['2025-05-12','วันวิสาขบูชา','national_holiday',2568],
      ['2025-05-30','วันอาสาฬหบูชา','national_holiday',2568],
      ['2025-07-11','วันเข้าพรรษา','national_holiday',2568],
      ['2025-07-28','วันเฉลิมพระชนมพรรษา ร.10','national_holiday',2568],
      ['2025-08-12','วันแม่แห่งชาติ','national_holiday',2568],
      ['2025-10-13','วันคล้ายวันสวรรคต ร.9','national_holiday',2568],
      ['2025-10-23','วันปิยมหาราช','national_holiday',2568],
      ['2025-12-05','วันพ่อแห่งชาติ','national_holiday',2568],
      ['2025-12-10','วันรัฐธรรมนูญ','national_holiday',2568],
      ['2025-12-31','วันสิ้นปี','national_holiday',2568],
    ];
    for (var hi = 0; hi < hol.length; hi++) {
      s.appendRow(hol[hi]);
    }
  }

  // --- 6. egp ---
  s = ss.getSheetByName('egp');
  if (!s) {
    s = ss.insertSheet('egp');
    s.getRange(1, 1, 1, 2).setValues([['project_no','project_name']]);
  }

  // --- 7. settings ---
  s = ss.getSheetByName('settings');
  if (!s) {
    s = ss.insertSheet('settings');
    s.getRange(1, 1, 1, 2).setValues([['key','value']]);
    s.appendRow(['buddhist_offset','543']);
    s.appendRow(['line_channel_access_token','']);
    s.appendRow(['line_notify_group_id','']);
    s.appendRow(['organization_name','ห้องยาสบปราบ']);
    s.appendRow(['auto_send_line','true']);
  }

  return 'สร้างชีททั้งหมดเรียบร้อยแล้ว';
}

// ============ API: AUTH ============

function apiLogin(fullname, password) {
  var users = getAllData('users');
  var user = null;
  for (var i = 0; i < users.length; i++) {
    if (String(users[i].fullname).trim() === String(fullname).trim() &&
        String(users[i].password) === String(password)) {
      user = users[i];
      break;
    }
  }
  if (user) {
    // Normalize is_admin to boolean
    user.is_admin = (user.is_admin === true || user.is_admin === 'true' || user.is_admin === 'TRUE');
    return { success: true, user: user };
  }
  return { success: false, message: 'ชื่อหรือรหัสผ่านไม่ถูกต้อง' };
}

// ============ API: USERS ============

function apiGetUsers() {
  return getAllData('users');
}

function apiAddUser(data) {
  if (!data.fullname || !data.password) {
    return { success: false, message: 'กรุณากรอกชื่อและรหัสผ่าน' };
  }
  // Check duplicate
  var users = getAllData('users');
  for (var i = 0; i < users.length; i++) {
    if (String(users[i].fullname).trim() === String(data.fullname).trim()) {
      return { success: false, message: 'ชื่อนี้มีอยู่ในระบบ' };
    }
  }

  data.user_id = generateId('USR');
  if (data.is_admin === undefined) data.is_admin = false;
  if (data.rate_per_hour === undefined) data.rate_per_hour = 0;
  addRow('users', data);
  return { success: true };
}

function apiUpdateUser(userId, data) {
  // Validate
  var users = getAllData('users');
  for (var i = 0; i < users.length; i++) {
    if (String(users[i].user_id) !== String(userId) &&
        String(users[i].fullname).trim() === String(data.fullname).trim()) {
      return { success: false, message: 'ชื่อนี้มีอยู่ในระบบ' };
    }
  }

  return { success: updateRow('users', 'user_id', userId, data) };
}

function apiDeleteUser(userId) {
  if (String(userId) === 'USR_001') {
    return { success: false, message: 'ไม่สามารถลบ Admin หลักได้' };
  }
  // Also delete related bookings
  var bookings = getAllData('bookings');
  for (var i = 0; i < bookings.length; i++) {
    if (String(bookings[i].user_id) === String(userId)) {
      deleteRow('bookings', 'booking_id', bookings[i].booking_id);
    }
  }
  // Delete related notifications
  var notifs = getAllData('notifications');
  for (var i = 0; i < notifs.length; i++) {
    if (String(notifs[i].user_id) === String(userId)) {
      deleteRow('notifications', 'notif_id', notifs[i].notif_id);
    }
  }
  return { success: deleteRow('users', 'user_id', userId) };
}

// ============ API: DUTY SHIFTS ============

function apiGetShifts() {
  return getAllData('duty_shifts');
}

function apiUpdateShift(shiftId, data) {
  return { success: updateRow('duty_shifts', 'shift_id', shiftId, data) };
}

// ============ API: BOOKINGS ============

function apiGetBookings(year, month) {
  var allBookings = getAllData('bookings');
  if (year && month) {
    var christianYear = parseInt(year) - BUDGET_OFFSET;
    var monthNum = parseInt(month);
    var prefix = christianYear + '-' + ('0' + monthNum).slice(-2);
    var result = [];
    for (var i = 0; i < allBookings.length; i++) {
      var dateStr = String(allBookings[i].duty_date).substring(0, 10);
      if (dateStr.substring(0, 7) === prefix && allBookings[i].status !== 'cancelled') {
        result.push(allBookings[i]);
      }
    }
    return result;
  }
  return allBookings.filter(function(b) { return b.status !== 'cancelled'; });
}

function apiCheckBooking(date) {
  var allBookings = getAllData('bookings');
  var dateStr = String(date).substring(0, 10);
  var existing = null;
  for (var i = 0; i < allBookings.length; i++) {
    var bDate = String(allBookings[i].duty_date).substring(0, 10);
    if (bDate === dateStr && allBookings[i].status !== 'cancelled') {
      existing = allBookings[i];
      break;
    }
  }
  if (existing) {
    return {
      booked: true,
      bookedBy: existing.booked_by,
      bookingId: existing.booking_id,
      shiftType: existing.shift_type
    };
  }
  return { booked: false };
}

function apiBookDuty(bookingData) {
  // Validate
  if (!bookingData.duty_date) {
    return { success: false, message: 'กรุณาระบุวันที่' };
  }

  // Check if already booked
  var check = apiCheckBooking(bookingData.duty_date);
  if (check.booked) {
    return {
      success: false,
      message: 'วันที่ ' + bookingData.duty_date + ' มีผู้จองแล้ว โดย ' + check.bookedBy
    };
  }

  bookingData.booking_id = generateId('BK');
  bookingData.booking_time = new Date().toISOString();
  bookingData.status = 'confirmed';
  bookingData.notes = bookingData.notes || '';

  addRow('bookings', bookingData);

  // Send LINE notification
  try {
    var settings = getAllData('settings');
    var autoSend = null;
    for (var s = 0; s < settings.length; s++) {
      if (String(settings[s].key) === 'auto_send_line') {
        autoSend = settings[s];
        break;
      }
    }
    if (autoSend && String(autoSend.value) === 'true') {
      sendLineNotification(bookingData);
    }
  } catch(e) {
    Logger.log('LINE notification error: ' + e.message);
  }

  return { success: true, bookingId: bookingData.booking_id };
}

function apiUpdateBooking(bookingId, data) {
  return { success: updateRow('bookings', 'booking_id', bookingId, data) };
}

function apiDeleteBooking(bookingId) {
  return { success: deleteRow('bookings', 'booking_id', bookingId) };
}

// ============ API: CALENDAR ============

function apiGetCalendar(year, month) {
  var christianYear = parseInt(year) - BUDGET_OFFSET;
  var monthNum = parseInt(month);
  var firstDay = new Date(christianYear, monthNum - 1, 1);
  var lastDay = new Date(christianYear, monthNum, 0);
  var daysInMonth = lastDay.getDate();
  var startDayOfWeek = firstDay.getDay(); // 0=Sun

  // Get holidays
  var holidays = apiGetMonthHolidays(year, month);
  var holidayMap = {};
  for (var h = 0; h < holidays.length; h++) {
    holidayMap[holidays[h].date] = holidays[h];
  }

  // Get bookings
  var bookings = apiGetBookings(year, month);
  var bookingMap = {};
  for (var b = 0; b < bookings.length; b++) {
    var bDateStr = String(bookings[b].duty_date).substring(0, 10);
    bookingMap[bDateStr] = bookings[b];
  }

  // Build calendar
  var days = [];
  var workingDays = 0;
  var weekendDays = 0;
  var nationalHolidayDays = 0;

  // Empty cells before first day
  for (var i = 0; i < startDayOfWeek; i++) {
    days.push(null);
  }

  for (var d = 1; d <= daysInMonth; d++) {
    var date = new Date(christianYear, monthNum - 1, d);
    var dateStr = formatDateKey(date);
    var dayOfWeek = date.getDay();
    var isWeekend = dayOfWeek === 0 || dayOfWeek === 6;
    var holiday = holidayMap[dateStr] || null;
    var booking = bookingMap[dateStr] || null;

    days.push({
      day: d,
      date: dateStr,
      dayOfWeek: dayOfWeek,
      dayNameThai: ['อาทิตย์','จันทร์','อังคาร','พุธ','พฤหัสบดี','ศุกร์','เสาร์'][dayOfWeek],
      isWeekend: isWeekend,
      isHoliday: !!holiday,
      holidayName: holiday ? holiday.name : '',
      holidayType: holiday ? holiday.type : '',
      booking: booking,
      isBooked: !!booking,
      bookedBy: booking ? booking.booked_by : '',
      shiftId: booking ? booking.shift_type : '',
      bookingId: booking ? booking.booking_id : ''
    });

    if (isWeekend) {
      weekendDays++;
    } else if (holiday && holiday.type === 'national_holiday') {
      nationalHolidayDays++;
    } else {
      workingDays++;
    }
  }

  return {
    year: year,
    month: month,
    thaiYear: parseInt(year),
    monthNameThai: getThaiMonthName(monthNum),
    days: days,
    workingDays: workingDays,
    weekendDays: weekendDays,
    nationalHolidayDays: nationalHolidayDays,
    totalDays: daysInMonth
  };
}

// ============ API: HOLIDAYS ============

function apiGetHolidays(year) {
  var allHolidays = getAllData('holidays');
  if (year) {
    return allHolidays.filter(function(h) { return String(h.year) === String(year); });
  }
  return allHolidays;
}

function apiAddHoliday(data) {
  if (!data.holiday_date || !data.holiday_name) {
    return { success: false, message: 'กรุณากรอกข้อมูลให้ครบ' };
  }
  addRow('holidays', data);
  return { success: true };
}

function apiDeleteHoliday(date) {
  return { success: deleteRow('holidays', 'holiday_date', date) };
}

function apiGetMonthHolidays(year, month) {
  var christianYear = parseInt(year) - BUDGET_OFFSET;
  var allHolidays = getAllData('holidays');
  var prefix = christianYear + '-' + ('0' + parseInt(month)).slice(-2);
  var daysInMonth = new Date(christianYear, parseInt(month), 0).getDate();
  var holidays = [];

  // Weekend holidays
  for (var d = 1; d <= daysInMonth; d++) {
    var date = new Date(christianYear, parseInt(month) - 1, d);
    var dayOfWeek = date.getDay();
    var dateStr = formatDateKey(date);
    if (dayOfWeek === 0 || dayOfWeek === 6) {
      holidays.push({
        date: dateStr,
        name: dayOfWeek === 0 ? 'วันอาทิตย์' : 'วันเสาร์',
        type: 'weekend'
      });
    }
  }

  // National holidays from sheet
  for (var h = 0; h < allHolidays.length; h++) {
    var hDate = String(allHolidays[h].holiday_date).substring(0, 10);
    if (hDate.substring(0, 7) === prefix && String(allHolidays[h].holiday_type) === 'national_holiday') {
      holidays.push({ date: hDate, name: allHolidays[h].holiday_name, type: 'national_holiday' });
    }
  }

  return holidays;
}

// ============ API: NOTIFICATIONS ============

function apiGetNotifications() {
  return getAllData('notifications');
}

function apiSetNotification(data) {
  var existing = getAllData('notifications');
  var match = null;
  for (var i = 0; i < existing.length; i++) {
    if (String(existing[i].user_id) === String(data.user_id) &&
        String(existing[i].shift_type) === String(data.shift_type)) {
      match = existing[i];
      break;
    }
  }
  if (match) {
    return { success: updateRow('notifications', 'notif_id', match.notif_id, { is_active: data.is_active }) };
  }
  data.notif_id = generateId('NOT');
  addRow('notifications', data);
  return { success: true };
}

// ============ API: EGP PROJECTS ============

function apiGetEgpProjects() {
  return getAllData('egp');
}

function apiAddEgpProject(data) {
  if (!data.project_no) return { success: false, message: 'กรุณาระบุเลขโครงการ' };
  addRow('egp', data);
  return { success: true };
}

function apiDeleteEgpProject(projectNo) {
  return { success: deleteRow('egp', 'project_no', projectNo) };
}

// ============ API: SETTINGS ============

function apiGetSettings() {
  return getAllData('settings');
}

function apiUpdateSetting(key, value) {
  var settings = getAllData('settings');
  var exists = null;
  for (var i = 0; i < settings.length; i++) {
    if (String(settings[i].key) === String(key)) {
      exists = settings[i];
      break;
    }
  }
  if (exists) {
    return { success: updateRow('settings', 'key', key, { value: value }) };
  }
  addRow('settings', { key: key, value: value });
  return { success: true };
}

// ============ API: ANALYSIS ============

function apiGetDocDutyUsers(year, month) {
  var bookings = apiGetBookings(year, month);
  // Filter doc duty bookings
  var docBookings = [];
  for (var i = 0; i < bookings.length; i++) {
    var st = String(bookings[i].shift_type);
    if (st === 'SHIFT_DOC_NORMAL' || st === 'SHIFT_DOC_WEEKEND') {
      docBookings.push(bookings[i]);
    }
  }

  // Get unique users with their dates
  var userMap = {};
  for (var i = 0; i < docBookings.length; i++) {
    var b = docBookings[i];
    var uid = String(b.user_id);
    if (!userMap[uid]) {
      userMap[uid] = {
        user_id: uid,
        booked_by: b.booked_by,
        dates: []
      };
    }
    userMap[uid].dates.push(String(b.duty_date).substring(0, 10));
  }

  var users = Object.keys(userMap).map(function(k) { return userMap[k]; });
  for (var u = 0; u < users.length; u++) {
    users[u].dates.sort();
  }
  return users;
}

function apiAnalyzeDocDuty(userId, year, month) {
  var bookings = apiGetBookings(year, month);
  var shifts = getAllData('duty_shifts');

  // Get user's doc duty bookings
  var userBookings = [];
  for (var i = 0; i < bookings.length; i++) {
    var st = String(bookings[i].shift_type);
    if (String(bookings[i].user_id) === String(userId) &&
        (st === 'SHIFT_DOC_NORMAL' || st === 'SHIFT_DOC_WEEKEND')) {
      userBookings.push(bookings[i]);
    }
  }

  if (userBookings.length === 0) {
    return { dates: [], total_projects: 0, projects_per_day: 0 };
  }

  // Sort by date
  userBookings.sort(function(a, b) { return String(a.duty_date).localeCompare(String(b.duty_date)); });

  // Get EGP projects
  var allEgp = getAllData('egp');
  var projects = [];
  for (var e = 0; e < allEgp.length; e++) {
    var pno = String(allEgp[e].project_no);
    if (pno && pno.trim() !== '') projects.push(pno);
  }
  var totalProjects = projects.length;

  // Distribution logic: 6 projects per day
  var result = [];
  for (var idx = 0; idx < userBookings.length; idx++) {
    var booking = userBookings[idx];
    var dateStr = String(booking.duty_date).substring(0, 10);
    var dateObj = parseDateString(dateStr);
    var dayOfWeek = dateObj ? dateObj.getDay() : -1;
    var isWeekend = dayOfWeek === 0 || dayOfWeek === 6;

    // Determine shift type
    var shiftId = isWeekend ? 'SHIFT_DOC_WEEKEND' : 'SHIFT_DOC_NORMAL';
    var shiftInfo = null;
    for (var s = 0; s < shifts.length; s++) {
      if (String(shifts[s].shift_id) === shiftId) {
        shiftInfo = shifts[s];
        break;
      }
    }

    // Distribute projects: 6 projects per day cycle
    var startIdx = idx * 6;
    var dayProjects = projects.slice(startIdx, startIdx + 6);

    result.push({
      date: dateStr,
      dateThai: dateObj ? formatDateNumeric(dateObj) : dateStr,
      dateThaiLong: dateObj ? formatDateThai(dateObj) : dateStr,
      dayOfWeek: ['อาทิตย์','จันทร์','อังคาร','พุธ','พฤหัสบดี','ศุกร์','เสาร์'][dayOfWeek],
      shift_id: shiftId,
      shift_name: shiftInfo ? shiftInfo.shift_name : shiftId,
      time_start: shiftInfo ? shiftInfo.time_start : '',
      time_end: shiftInfo ? shiftInfo.time_end : '',
      hours: shiftInfo ? parseInt(shiftInfo.hours) : 0,
      projects: dayProjects,
      project_count: dayProjects.length
    });
  }

  var totalHours = 0;
  var totalProjectWork = 0;
  for (var r = 0; r < result.length; r++) {
    totalHours += result[r].hours;
    totalProjectWork += result[r].project_count;
  }

  return {
    dates: result,
    total_projects: totalProjects,
    projects_per_day: 6,
    total_hours: totalHours,
    total_project_work: totalProjectWork
  };
}

// ============ API: REPORTS ============

function apiGetMonthlyReport(year, month) {
  var bookings = apiGetBookings(year, month);
  var users = getAllData('users');
  var shifts = getAllData('duty_shifts');

  var reportBookings = [];
  for (var i = 0; i < bookings.length; i++) {
    var b = bookings[i];
    var shift = null;
    for (var s = 0; s < shifts.length; s++) {
      if (String(shifts[s].shift_id) === String(b.shift_type)) {
        shift = shifts[s];
        break;
      }
    }
    reportBookings.push({
      booking_id: b.booking_id,
      duty_date: b.duty_date,
      shift_type: b.shift_type,
      user_id: b.user_id,
      booked_by: b.booked_by,
      shift_name: shift ? shift.shift_name : String(b.shift_type),
      shift_type_name: shift ? shift.shift_type : '',
      time_start: shift ? shift.time_start : '',
      time_end: shift ? shift.time_end : '',
      hours: shift ? shift.hours : 0,
      notes: b.notes || ''
    });
  }

  return {
    year: year,
    month: month,
    thaiYear: parseInt(year),
    bookings: reportBookings
  };
}

function apiGetPaymentReport(year, month) {
  var bookings = apiGetBookings(year, month);
  var users = getAllData('users');
  var shifts = getAllData('duty_shifts');

  // Group by user
  var userSummary = {};
  for (var i = 0; i < bookings.length; i++) {
    var b = bookings[i];
    var uid = String(b.user_id);
    if (!userSummary[uid]) {
      userSummary[uid] = {
        user_id: uid,
        booked_by: b.booked_by,
        normal_days: 0,
        weekend_days: 0,
        doc_normal_days: 0,
        doc_weekend_days: 0,
        total_hours: 0,
        total_payment: 0,
        details: []
      };
    }

    var shift = null;
    for (var s = 0; s < shifts.length; s++) {
      if (String(shifts[s].shift_id) === String(b.shift_type)) {
        shift = shifts[s];
        break;
      }
    }
    var hours = shift ? parseInt(shift.hours) : 0;
    var multiplier = shift ? parseFloat(shift.rate_multiplier) : 1;

    var user = null;
    for (var u = 0; u < users.length; u++) {
      if (String(users[u].user_id) === uid) {
        user = users[u];
        break;
      }
    }
    var ratePerHour = user ? parseFloat(user.rate_per_hour) || 0 : 0;
    var payment = hours * ratePerHour * multiplier;

    userSummary[uid].total_hours += hours;
    userSummary[uid].total_payment += payment;

    // Count by type
    if (String(b.shift_type) === 'SHIFT_NORMAL') userSummary[uid].normal_days++;
    else if (String(b.shift_type) === 'SHIFT_WEEKEND') userSummary[uid].weekend_days++;
    else if (String(b.shift_type) === 'SHIFT_DOC_NORMAL') userSummary[uid].doc_normal_days++;
    else if (String(b.shift_type) === 'SHIFT_DOC_WEEKEND') userSummary[uid].doc_weekend_days++;

    userSummary[uid].details.push({
      date: b.duty_date,
      shift_type: b.shift_type,
      shift_name: shift ? shift.shift_name : '',
      hours: hours,
      rate: ratePerHour,
      multiplier: multiplier,
      payment: payment
    });
  }

  return {
    year: year,
    month: month,
    thaiYear: parseInt(year),
    summary: Object.keys(userSummary).map(function(k) { return userSummary[k]; })
  };
}

function apiGetMyDuty(userId, year, month) {
  var bookings = apiGetBookings(year, month);
  var shifts = getAllData('duty_shifts');

  var myBookings = [];
  for (var i = 0; i < bookings.length; i++) {
    if (String(bookings[i].user_id) === String(userId)) {
      myBookings.push(bookings[i]);
    }
  }

  var normalCount = 0, weekendCount = 0, docNormalCount = 0, docWeekendCount = 0, totalHours = 0;
  var reportBookings = [];

  for (var i = 0; i < myBookings.length; i++) {
    var b = myBookings[i];
    var shift = null;
    for (var s = 0; s < shifts.length; s++) {
      if (String(shifts[s].shift_id) === String(b.shift_type)) {
        shift = shifts[s];
        break;
      }
    }

    if (String(b.shift_type) === 'SHIFT_NORMAL') normalCount++;
    else if (String(b.shift_type) === 'SHIFT_WEEKEND') weekendCount++;
    else if (String(b.shift_type) === 'SHIFT_DOC_NORMAL') docNormalCount++;
    else if (String(b.shift_type) === 'SHIFT_DOC_WEEKEND') docWeekendCount++;

    if (shift) totalHours += parseInt(shift.hours);

    reportBookings.push({
      booking_id: b.booking_id,
      duty_date: b.duty_date,
      shift_type: b.shift_type,
      booked_by: b.booked_by,
      shift_name: shift ? shift.shift_name : String(b.shift_type),
      shift_type_name: shift ? shift.shift_type : '',
      hours: shift ? parseInt(shift.hours) : 0,
      time_start: shift ? shift.time_start : '',
      time_end: shift ? shift.time_end : '',
      notes: b.notes || ''
    });
  }

  return {
    user_id: userId,
    bookings: reportBookings,
    normal_days: normalCount,
    weekend_days: weekendCount,
    doc_normal_days: docNormalCount,
    doc_weekend_days: docWeekendCount,
    total_days: myBookings.length,
    total_hours: totalHours
  };
}

// ============ API: DOC DUTY ATTACHMENT (เอกสารแนบเวรเอกสาร) ============

function apiGetDocDutyAttachment(year, month) {
  var bookings = apiGetBookings(year, month);
  var shifts = getAllData('duty_shifts');

  // Filter only doc duty
  var docBookings = [];
  for (var i = 0; i < bookings.length; i++) {
    var st = String(bookings[i].shift_type);
    if (st === 'SHIFT_DOC_NORMAL' || st === 'SHIFT_DOC_WEEKEND') {
      docBookings.push(bookings[i]);
    }
  }

  // Get unique users
  var userMap = {};
  for (var i = 0; i < docBookings.length; i++) {
    var b = docBookings[i];
    var uid = String(b.user_id);
    if (!userMap[uid]) {
      userMap[uid] = {
        user_id: uid,
        booked_by: b.booked_by,
        dates: []
      };
    }
    userMap[uid].dates.push(String(b.duty_date).substring(0, 10));
  }

  // Get EGP projects
  var allEgp = getAllData('egp');
  var projects = [];
  for (var e = 0; e < allEgp.length; e++) {
    var pno = String(allEgp[e].project_no);
    if (pno && pno.trim() !== '') projects.push(pno);
  }
  var totalProjects = projects.length;

  // Analyze each user
  var userKeys = Object.keys(userMap);
  var result = [];
  for (var u = 0; u < userKeys.length; u++) {
    var user = userMap[userKeys[u]];
    user.dates.sort();
    var numDays = user.dates.length;

    var dayDetails = [];
    for (var idx = 0; idx < user.dates.length; idx++) {
      var dateStr = user.dates[idx];
      var dateObj = parseDateString(dateStr);
      var dayOfWeek = dateObj ? dateObj.getDay() : -1;
      var isWeekend = dayOfWeek === 0 || dayOfWeek === 6;
      var shiftId = isWeekend ? 'SHIFT_DOC_WEEKEND' : 'SHIFT_DOC_NORMAL';
      var shiftInfo = null;
      for (var s = 0; s < shifts.length; s++) {
        if (String(shifts[s].shift_id) === shiftId) {
          shiftInfo = shifts[s];
          break;
        }
      }

      var startIdx = idx * 6;
      var dayProjects = projects.slice(startIdx, startIdx + 6);

      dayDetails.push({
        date: dateStr,
        dateThai: dateObj ? formatDateNumeric(dateObj) : dateStr,
        dayName: ['อาทิตย์','จันทร์','อังคาร','พุธ','พฤหัสบดี','ศุกร์','เสาร์'][dayOfWeek],
        shift_name: shiftInfo ? shiftInfo.shift_name : '',
        time_start: shiftInfo ? shiftInfo.time_start : '',
        time_end: shiftInfo ? shiftInfo.time_end : '',
        hours: shiftInfo ? parseInt(shiftInfo.hours) : 0,
        projects: dayProjects,
        project_count: dayProjects.length
      });
    }

    var dayTotalHours = 0;
    var dayTotalProjects = 0;
    for (var dd = 0; dd < dayDetails.length; dd++) {
      dayTotalHours += dayDetails[dd].hours;
      dayTotalProjects += dayDetails[dd].project_count;
    }

    result.push({
      user_id: user.user_id,
      booked_by: user.booked_by,
      total_days: numDays,
      day_details: dayDetails,
      total_hours: dayTotalHours,
      total_projects: dayTotalProjects
    });
  }

  return {
    year: year,
    month: month,
    thaiYear: parseInt(year),
    total_projects_available: totalProjects,
    users: result
  };
}

// ============ LINE NOTIFICATION ============

function sendLineNotification(bookingData) {
  var settings = getAllData('settings');
  var token = '';
  var targetId = '';
  for (var i = 0; i < settings.length; i++) {
    if (String(settings[i].key) === 'line_channel_access_token') token = settings[i].value;
    if (String(settings[i].key) === 'line_notify_group_id') targetId = settings[i].value;
  }

  if (!token || !targetId) {
    Logger.log('LINE: Token or target ID not configured');
    return;
  }

  var shifts = getAllData('duty_shifts');
  var shift = null;
  for (var s = 0; s < shifts.length; s++) {
    if (String(shifts[s].shift_id) === String(bookingData.shift_type)) {
      shift = shifts[s];
      break;
    }
  }
  var shiftName = shift ? shift.shift_name : String(bookingData.shift_type);
  var shiftTime = shift ? shift.time_start + ' - ' + shift.time_end : '';

  var dateObj = parseDateString(bookingData.duty_date);
  var dateThai = dateObj ? formatDateThai(dateObj) : String(bookingData.duty_date);

  var message = '📋 แจ้งเตือนการจองเวร - ห้องยาสบปราบ\n' +
    '\n' +
    '👤 ผู้จอง: ' + bookingData.booked_by + '\n' +
    '📅 วันที่: ' + dateThai + '\n' +
    '🔄 ประเภทเวร: ' + shiftName + '\n' +
    '⏰ เวลา: ' + shiftTime;

  try {
    var url = 'https://api.line.me/v2/bot/message/push';
    var options = {
      method: 'post',
      headers: {
        'Content-Type': 'application/json',
        'Authorization': 'Bearer ' + token
      },
      payload: JSON.stringify({
        to: targetId,
        messages: [{ type: 'text', text: message }]
      }),
      muteHttpExceptions: true
    };
    var response = UrlFetchApp.fetch(url, options);
    Logger.log('LINE response: ' + response.getResponseCode());
  } catch(e) {
    Logger.log('LINE API Error: ' + e.message);
  }
}

// ============ WEB APP SERVING ============

function doGet(e) {
  return HtmlService.createTemplateFromFile('index')
    .evaluate()
    .setTitle('ระบบจัดตารางเวร - ห้องยาสบปราบ')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1, maximum-scale=1, user-scalable=no')
    .addMetaTag('apple-mobile-web-app-capable', 'yes')
    .addMetaTag('apple-mobile-web-app-status-bar-style', 'black-translucent')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function doPost(e) {
  try {
    var params = JSON.parse(e.postData.contents);
    var action = params.action;
    var data = params.data || {};

    var result;
    switch(action) {
      case 'login': result = apiLogin(data.fullname, data.password); break;
      case 'getUsers': result = apiGetUsers(); break;
      case 'addUser': result = apiAddUser(data); break;
      case 'updateUser': result = apiUpdateUser(data.user_id, data); break;
      case 'deleteUser': result = apiDeleteUser(data.user_id); break;
      case 'getShifts': result = apiGetShifts(); break;
      case 'updateShift': result = apiUpdateShift(data.shift_id, data); break;
      case 'getBookings': result = apiGetBookings(data.year, data.month); break;
      case 'checkBooking': result = apiCheckBooking(data.duty_date); break;
      case 'bookDuty': result = apiBookDuty(data); break;
      case 'updateBooking': result = apiUpdateBooking(data.booking_id, data); break;
      case 'deleteBooking': result = apiDeleteBooking(data.booking_id); break;
      case 'getCalendar': result = apiGetCalendar(data.year, data.month); break;
      case 'getHolidays': result = apiGetHolidays(data.year); break;
      case 'getMonthHolidays': result = apiGetMonthHolidays(data.year, data.month); break;
      case 'addHoliday': result = apiAddHoliday(data); break;
      case 'deleteHoliday': result = apiDeleteHoliday(data.holiday_date); break;
      case 'getNotifications': result = apiGetNotifications(); break;
      case 'setNotification': result = apiSetNotification(data); break;
      case 'getEgpProjects': result = apiGetEgpProjects(); break;
      case 'addEgpProject': result = apiAddEgpProject(data); break;
      case 'deleteEgpProject': result = apiDeleteEgpProject(data.project_no); break;
      case 'getSettings': result = apiGetSettings(); break;
      case 'updateSetting': result = apiUpdateSetting(data.key, data.value); break;
      case 'getDocDutyUsers': result = apiGetDocDutyUsers(data.year, data.month); break;
      case 'analyzeDocDuty': result = apiAnalyzeDocDuty(data.user_id, data.year, data.month); break;
      case 'getDocDutyAttachment': result = apiGetDocDutyAttachment(data.year, data.month); break;
      case 'getMonthlyReport': result = apiGetMonthlyReport(data.year, data.month); break;
      case 'getPaymentReport': result = apiGetPaymentReport(data.year, data.month); break;
      case 'getMyDuty': result = apiGetMyDuty(data.user_id, data.year, data.month); break;
      default: result = { error: 'Unknown action: ' + action };
    }

    return ContentService.createTextOutput(JSON.stringify(result))
      .setMimeType(ContentService.MimeType.JSON);
  } catch(err) {
    return ContentService.createTextOutput(JSON.stringify({ error: err.message }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}
