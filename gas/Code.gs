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
const BUDGET_OFFSET = 543; // พ.ศ. - ค.ศ.

// ============ SHEET HELPERS ============

function getSheet(name) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(name);
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
    const sheet = getSheet(sheetName);
    const lastRow = sheet.getLastRow();
    const lastCol = sheet.getLastColumn();
    if (lastRow <= 1 || lastCol === 0) return [];
    const data = sheet.getRange(1, 1, lastRow, lastCol).getValues();
    if (data.length <= 1) return [];
    const headers = data[0];
    return data.slice(1).map(row => {
      const obj = {};
      headers.forEach((h, i) => { obj[h] = row[i]; });
      return obj;
    });
  } catch(e) {
    Logger.log('getAllData error [' + sheetName + ']: ' + e.message);
    return [];
  }
}

/**
 * เพิ่มแถวใหม่ในชีท
 */
function addRow(sheetName, rowData) {
  const sheet = getSheet(sheetName);
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const row = headers.map(h => {
    const val = rowData[h];
    if (val === undefined || val === null) return '';
    return val;
  });
  sheet.appendRow(row);
  return true;
}

/**
 * อัพเดทแถวตามเงื่อนไข
 */
function updateRow(sheetName, matchField, matchValue, updateData) {
  const sheet = getSheet(sheetName);
  const data = sheet.getDataRange().getValues();
  if (data.length <= 1) return false;
  const headers = data[0];
  const matchIdx = headers.indexOf(matchField);
  if (matchIdx === -1) return false;

  const updateKeys = Object.keys(updateData);
  const updateCols = updateKeys.map(k => headers.indexOf(k));

  for (let i = 1; i < data.length; i++) {
    if (String(data[i][matchIdx]).trim() === String(matchValue).trim()) {
      updateCols.forEach((colIdx, idx) => {
        if (colIdx >= 0) {
          sheet.getRange(i + 1, colIdx + 1).setValue(updateData[updateKeys[idx]]);
        }
      });
      return true;
    }
  }
  return false;
}

/**
 * ลบแถวตามเงื่อนไข
 */
function deleteRow(sheetName, matchField, matchValue) {
  const sheet = getSheet(sheetName);
  const data = sheet.getDataRange().getValues();
  if (data.length <= 1) return false;
  const headers = data[0];
  const matchIdx = headers.indexOf(matchField);
  if (matchIdx === -1) return false;

  for (let i = data.length - 1; i >= 1; i--) {
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
  const y = date.getFullYear();
  const m = String(date.getMonth() + 1).padStart(2, '0');
  const d = String(date.getDate()).padStart(2, '0');
  return y + '-' + m + '-' + d;
}

/**
 * แปลง Date object เป็นรูปแบบไทย วัน/เดือน/ปี พ.ศ.
 */
function formatDateThai(date) {
  if (!(date instanceof Date) || isNaN(date)) return '';
  const thaiMonths = ['ม.ค.','ก.พ.','มี.ค.','เม.ย.','พ.ค.','มิ.ย.','ก.ค.','ส.ค.','ก.ย.','ต.ค.','พ.ย.','ธ.ค.'];
  return date.getDate() + ' ' + thaiMonths[date.getMonth()] + ' พ.ศ. ' + (date.getFullYear() + BUDGET_OFFSET);
}

/**
 * แปลง Date object เป็นรูปแบบ ว/ด/พ.ศ.
 */
function formatDateNumeric(date) {
  if (!(date instanceof Date) || isNaN(date)) return '';
  const d = String(date.getDate()).padStart(2, '0');
  const m = String(date.getMonth() + 1).padStart(2, '0');
  const y = date.getFullYear() + BUDGET_OFFSET;
  return d + '/' + m + '/' + y;
}

/**
 * แปลง YYYY-MM-DD เป็น Date object
 */
function parseDateString(dateStr) {
  if (dateStr instanceof Date) return dateStr;
  if (!dateStr) return null;
  const str = String(dateStr);
  const parts = str.substring(0, 10).split('-');
  if (parts.length !== 3) return null;
  return new Date(parseInt(parts[0]), parseInt(parts[1]) - 1, parseInt(parts[2]));
}

/**
 * ชื่อเดือนไทย
 */
function getThaiMonthName(monthNum) {
  const months = ['','มกราคม','กุมภาพันธ์','มีนาคม','เมษายน','พฤษภาคม','มิถุนายน',
    'กรกฎาคม','สิงหาคม','กันยายน','ตุลาคม','พฤศจิกายน','ธันวาคม'];
  return months[parseInt(monthNum)] || '';
}

/**
 * ชื่อเดือนไทยแบบย่อ
 */
function getThaiMonthNameShort(monthIdx) {
  const names = ['ม.ค.','ก.พ.','มี.ค.','เม.ย.','พ.ค.','มิ.ย.','ก.ค.','ส.ค.','ก.ย.','ต.ค.','พ.ย.','ธ.ค.'];
  return names[parseInt(monthIdx)] || '';
}

// ============ INITIALIZE SHEETS ============

function initializeSheets() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  // --- 1. users ---
  let s = ss.getSheetByName('users');
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
    const hol = [
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
    hol.forEach(h => s.appendRow(h));
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
  const users = getAllData('users');
  const user = users.find(u =>
    String(u.fullname).trim() === String(fullname).trim() &&
    String(u.password) === String(password)
  );
  if (user) {
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
  const users = getAllData('users');
  const dup = users.find(u => String(u.fullname).trim() === String(data.fullname).trim());
  if (dup) return { success: false, message: 'ชื่อนี้มีอยู่แล้วในระบบ' };

  data.user_id = generateId('USR');
  if (data.is_admin === undefined) data.is_admin = false;
  if (data.rate_per_hour === undefined) data.rate_per_hour = 0;
  addRow('users', data);
  return { success: true };
}

function apiUpdateUser(userId, data) {
  // Validate
  const users = getAllData('users');
  const dup = users.find(u =>
    String(u.user_id) !== String(userId) &&
    String(u.fullname).trim() === String(data.fullname).trim()
  );
  if (dup) return { success: false, message: 'ชื่อนี้มีอยู่แล้วในระบบ' };

  return { success: updateRow('users', 'user_id', userId, data) };
}

function apiDeleteUser(userId) {
  if (String(userId) === 'USR_001') {
    return { success: false, message: 'ไม่สามารถลบ Admin หลักได้' };
  }
  // Also delete related bookings
  const bookings = getAllData('bookings');
  bookings.filter(b => String(b.user_id) === String(userId)).forEach(b => {
    deleteRow('bookings', 'booking_id', b.booking_id);
  });
  // Delete related notifications
  const notifs = getAllData('notifications');
  notifs.filter(n => String(n.user_id) === String(userId)).forEach(n => {
    deleteRow('notifications', 'notif_id', n.notif_id);
  });
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
  const allBookings = getAllData('bookings');
  if (year && month) {
    const christianYear = parseInt(year) - BUDGET_OFFSET;
    const monthNum = parseInt(month);
    const prefix = christianYear + '-' + String(monthNum).padStart(2, '0');
    return allBookings.filter(b => {
      const dateStr = String(b.duty_date).substring(0, 10);
      return dateStr.startsWith(prefix) && b.status !== 'cancelled';
    });
  }
  return allBookings.filter(b => b.status !== 'cancelled');
}

function apiCheckBooking(date) {
  const allBookings = getAllData('bookings');
  const dateStr = String(date).substring(0, 10);
  const existing = allBookings.find(b => {
    const bDate = String(b.duty_date).substring(0, 10);
    return bDate === dateStr && b.status !== 'cancelled';
  });
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
  const check = apiCheckBooking(bookingData.duty_date);
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
    const settings = getAllData('settings');
    const autoSend = settings.find(s => s.key === 'auto_send_line');
    if (autoSend && autoSend.value === 'true') {
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
  const christianYear = parseInt(year) - BUDGET_OFFSET;
  const monthNum = parseInt(month);
  const firstDay = new Date(christianYear, monthNum - 1, 1);
  const lastDay = new Date(christianYear, monthNum, 0);
  const daysInMonth = lastDay.getDate();
  const startDayOfWeek = firstDay.getDay(); // 0=Sun

  // Get holidays
  const holidays = apiGetMonthHolidays(year, month);
  const holidayMap = {};
  holidays.forEach(h => { holidayMap[h.date] = h; });

  // Get bookings
  const bookings = apiGetBookings(year, month);
  const bookingMap = {};
  bookings.forEach(b => {
    const dateStr = String(b.duty_date).substring(0, 10);
    bookingMap[dateStr] = b;
  });

  // Build calendar
  const days = [];
  let workingDays = 0;
  let weekendDays = 0;
  let holidayDays = 0;

  // Empty cells before first day
  for (let i = 0; i < startDayOfWeek; i++) {
    days.push(null);
  }

  for (let d = 1; d <= daysInMonth; d++) {
    const date = new Date(christianYear, monthNum - 1, d);
    const dateStr = formatDateKey(date);
    const dayOfWeek = date.getDay();
    const isWeekend = dayOfWeek === 0 || dayOfWeek === 6;
    const holiday = holidayMap[dateStr];
    const booking = bookingMap[dateStr];

    days.push({
      day: d,
      date: dateStr,
      dayOfWeek: dayOfWeek,
      dayNameThai: ['อาทิตย์','จันทร์','อังคาร','พุธ','พฤหัสบดี','ศุกร์','เสาร์'][dayOfWeek],
      isWeekend: isWeekend,
      isHoliday: !!holiday,
      holidayName: holiday ? holiday.name : '',
      holidayType: holiday ? holiday.type : '',
      booking: booking || null,
      isBooked: !!booking,
      bookedBy: booking ? booking.booked_by : '',
      shiftId: booking ? booking.shift_type : '',
      bookingId: booking ? booking.booking_id : ''
    });

    if (isWeekend || holiday) {
      holidayDays++;
      if (isWeekend) weekendDays++;
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
    holidayDays: holidayDays,
    totalDays: daysInMonth
  };
}

// ============ API: HOLIDAYS ============

function apiGetHolidays(year) {
  const allHolidays = getAllData('holidays');
  if (year) {
    return allHolidays.filter(h => String(h.year) === String(year));
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
  const christianYear = parseInt(year) - BUDGET_OFFSET;
  const allHolidays = getAllData('holidays');
  const prefix = christianYear + '-' + String(parseInt(month)).padStart(2, '0');
  const daysInMonth = new Date(christianYear, parseInt(month), 0).getDate();
  const holidays = [];

  // Weekend holidays
  for (let d = 1; d <= daysInMonth; d++) {
    const date = new Date(christianYear, parseInt(month) - 1, d);
    const dayOfWeek = date.getDay();
    const dateStr = formatDateKey(date);
    if (dayOfWeek === 0 || dayOfWeek === 6) {
      holidays.push({
        date: dateStr,
        name: dayOfWeek === 0 ? 'วันอาทิตย์' : 'วันเสาร์',
        type: 'weekend'
      });
    }
  }

  // National holidays from sheet
  allHolidays.forEach(h => {
    const hDate = String(h.holiday_date).substring(0, 10);
    if (hDate.startsWith(prefix) && h.holiday_type === 'national_holiday') {
      holidays.push({ date: hDate, name: h.holiday_name, type: 'national_holiday' });
    }
  });

  return holidays;
}

// ============ API: NOTIFICATIONS ============

function apiGetNotifications() {
  return getAllData('notifications');
}

function apiSetNotification(data) {
  const existing = getAllData('notifications');
  const match = existing.find(n =>
    String(n.user_id) === String(data.user_id) &&
    String(n.shift_type) === String(data.shift_type)
  );
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
  const settings = getAllData('settings');
  const exists = settings.find(s => String(s.key) === String(key));
  if (exists) {
    return { success: updateRow('settings', 'key', key, { value: value }) };
  }
  addRow('settings', { key: key, value: value });
  return { success: true };
}

// ============ API: ANALYSIS ============

function apiGetDocDutyUsers(year, month) {
  const bookings = apiGetBookings(year, month);
  // Filter doc duty bookings
  const docBookings = bookings.filter(b => {
    const st = String(b.shift_type);
    return st === 'SHIFT_DOC_NORMAL' || st === 'SHIFT_DOC_WEEKEND';
  });

  // Get unique users with their dates
  const userMap = {};
  docBookings.forEach(b => {
    const uid = String(b.user_id);
    if (!userMap[uid]) {
      userMap[uid] = {
        user_id: uid,
        booked_by: b.booked_by,
        dates: []
      };
    }
    userMap[uid].dates.push(String(b.duty_date).substring(0, 10));
  });

  const users = Object.values(userMap);
  users.forEach(u => u.dates.sort());
  return users;
}

function apiAnalyzeDocDuty(userId, year, month) {
  const bookings = apiGetBookings(year, month);
  const shifts = getAllData('duty_shifts');

  // Get user's doc duty bookings
  const userBookings = bookings.filter(b => {
    const st = String(b.shift_type);
    return String(b.user_id) === String(userId) &&
      (st === 'SHIFT_DOC_NORMAL' || st === 'SHIFT_DOC_WEEKEND');
  });

  if (userBookings.length === 0) {
    return { dates: [], total_projects: 0, projects_per_day: 0 };
  }

  // Sort by date
  userBookings.sort((a, b) => String(a.duty_date).localeCompare(String(b.duty_date)));

  // Get EGP projects
  const allEgp = getAllData('egp');
  const projects = allEgp.map(p => String(p.project_no)).filter(Boolean);
  const totalProjects = projects.length;

  // Distribution logic: total / 6 (since 1 person has 6 days)
  const numDays = userBookings.length;
  const projectsPerDay = numDays > 0 ? Math.ceil(totalProjects / numDays) : 0;

  const result = userBookings.map((booking, idx) => {
    const dateStr = String(booking.duty_date).substring(0, 10);
    const dateObj = parseDateString(dateStr);
    const dayOfWeek = dateObj ? dateObj.getDay() : -1;
    const isWeekend = dayOfWeek === 0 || dayOfWeek === 6;

    // Determine shift type
    const shiftId = isWeekend ? 'SHIFT_DOC_WEEKEND' : 'SHIFT_DOC_NORMAL';
    const shiftInfo = shifts.find(s => String(s.shift_id) === shiftId);

    // Distribute projects: 6 projects per day cycle
    const startIdx = idx * 6;
    const dayProjects = projects.slice(startIdx, startIdx + 6);

    return {
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
    };
  });

  return {
    dates: result,
    total_projects: totalProjects,
    projects_per_day: 6,
    total_hours: result.reduce((s, d) => s + d.hours, 0),
    total_project_work: result.reduce((s, d) => s + d.project_count, 0)
  };
}

// ============ API: REPORTS ============

function apiGetMonthlyReport(year, month) {
  const bookings = apiGetBookings(year, month);
  const users = getAllData('users');
  const shifts = getAllData('duty_shifts');

  return {
    year: year,
    month: month,
    thaiYear: parseInt(year),
    bookings: bookings.map(b => {
      const shift = shifts.find(s => String(s.shift_id) === String(b.shift_type));
      return {
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
      };
    })
  };
}

function apiGetPaymentReport(year, month) {
  const bookings = apiGetBookings(year, month);
  const users = getAllData('users');
  const shifts = getAllData('duty_shifts');

  // Group by user
  const userSummary = {};
  bookings.forEach(b => {
    const uid = String(b.user_id);
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

    const shift = shifts.find(s => String(s.shift_id) === String(b.shift_type));
    const hours = shift ? parseInt(shift.hours) : 0;
    const multiplier = shift ? parseFloat(shift.rate_multiplier) : 1;

    const user = users.find(u => String(u.user_id) === uid);
    const ratePerHour = user ? parseFloat(user.rate_per_hour) || 0 : 0;
    const payment = hours * ratePerHour * multiplier;

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
  });

  return {
    year: year,
    month: month,
    thaiYear: parseInt(year),
    summary: Object.values(userSummary)
  };
}

function apiGetMyDuty(userId, year, month) {
  const bookings = apiGetBookings(year, month);
  const shifts = getAllData('duty_shifts');

  const myBookings = bookings.filter(b => String(b.user_id) === String(userId));

  return {
    user_id: userId,
    bookings: myBookings.map(b => {
      const shift = shifts.find(s => String(s.shift_id) === String(b.shift_type));
      return {
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
      };
    }),
    normal_days: myBookings.filter(b => String(b.shift_type) === 'SHIFT_NORMAL'),
    weekend_days: myBookings.filter(b => String(b.shift_type) === 'SHIFT_WEEKEND'),
    doc_normal_days: myBookings.filter(b => String(b.shift_type) === 'SHIFT_DOC_NORMAL'),
    doc_weekend_days: myBookings.filter(b => String(b.shift_type) === 'SHIFT_DOC_WEEKEND'),
    total_days: myBookings.length,
    total_hours: myBookings.reduce((sum, b) => {
      const shift = shifts.find(s => String(s.shift_id) === String(b.shift_type));
      return sum + (shift ? parseInt(shift.hours) : 0);
    }, 0)
  };
}

// ============ API: DOC DUTY ATTACHMENT (เอกสารแนบเวรเอกสาร) ============

function apiGetDocDutyAttachment(year, month) {
  const bookings = apiGetBookings(year, month);
  const shifts = getAllData('duty_shifts');

  // Filter only doc duty
  const docBookings = bookings.filter(b => {
    const st = String(b.shift_type);
    return st === 'SHIFT_DOC_NORMAL' || st === 'SHIFT_DOC_WEEKEND';
  });

  // Get unique users
  const userMap = {};
  docBookings.forEach(b => {
    const uid = String(b.user_id);
    if (!userMap[uid]) {
      userMap[uid] = {
        user_id: uid,
        booked_by: b.booked_by,
        dates: []
      };
    }
    userMap[uid].dates.push(String(b.duty_date).substring(0, 10));
  });

  // Get EGP projects
  const allEgp = getAllData('egp');
  const projects = allEgp.map(p => String(p.project_no)).filter(Boolean);
  const totalProjects = projects.length;

  // Analyze each user
  const users = Object.values(userMap);
  const result = users.map(user => {
    user.dates.sort();
    const numDays = user.dates.length;
    const projectsPerDay = numDays > 0 ? Math.ceil(totalProjects / numDays) : 0;

    const dayDetails = user.dates.map((dateStr, idx) => {
      const dateObj = parseDateString(dateStr);
      const dayOfWeek = dateObj ? dateObj.getDay() : -1;
      const isWeekend = dayOfWeek === 0 || dayOfWeek === 6;
      const shiftId = isWeekend ? 'SHIFT_DOC_WEEKEND' : 'SHIFT_DOC_NORMAL';
      const shiftInfo = shifts.find(s => String(s.shift_id) === shiftId);

      const startIdx = idx * 6;
      const dayProjects = projects.slice(startIdx, startIdx + 6);

      return {
        date: dateStr,
        dateThai: dateObj ? formatDateNumeric(dateObj) : dateStr,
        dayName: ['อาทิตย์','จันทร์','อังคาร','พุธ','พฤหัสบดี','ศุกร์','เสาร์'][dayOfWeek],
        shift_name: shiftInfo ? shiftInfo.shift_name : '',
        time_start: shiftInfo ? shiftInfo.time_start : '',
        time_end: shiftInfo ? shiftInfo.time_end : '',
        hours: shiftInfo ? parseInt(shiftInfo.hours) : 0,
        projects: dayProjects,
        project_count: dayProjects.length
      };
    });

    return {
      user_id: user.user_id,
      booked_by: user.booked_by,
      total_days: numDays,
      day_details: dayDetails,
      total_hours: dayDetails.reduce((s, d) => s + d.hours, 0),
      total_projects: dayDetails.reduce((s, d) => s + d.project_count, 0)
    };
  });

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
  const settings = getAllData('settings');
  const tokenSetting = settings.find(s => String(s.key) === 'line_channel_access_token');
  const token = tokenSetting ? tokenSetting.value : '';
  const groupId = settings.find(s => String(s.key) === 'line_notify_group_id');
  const targetId = groupId ? groupId.value : '';

  if (!token || !targetId) {
    Logger.log('LINE: Token or target ID not configured');
    return;
  }

  const shifts = getAllData('duty_shifts');
  const shift = shifts.find(s => String(s.shift_id) === String(bookingData.shift_type));
  const shiftName = shift ? shift.shift_name : String(bookingData.shift_type);
  const shiftTime = shift ? shift.time_start + ' - ' + shift.time_end : '';

  const dateObj = parseDateString(bookingData.duty_date);
  const dateThai = dateObj ? formatDateThai(dateObj) : String(bookingData.duty_date);

  const message = [
    '📋 แจ้งเตือนการจองเวร - ห้องยาสบปราบ',
    '',
    '👤 ผู้จอง: ' + bookingData.booked_by,
    '📅 วันที่: ' + dateThai,
    '🔄 ประเภทเวร: ' + shiftName,
    '⏰ เวลา: ' + shiftTime,
    ''
  ].join('\n');

  try {
    const url = 'https://api.line.me/v2/bot/message/push';
    const options = {
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
    const response = UrlFetchApp.fetch(url, options);
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
    const params = JSON.parse(e.postData.contents);
    const action = params.action;
    const data = params.data || {};

    let result;
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
