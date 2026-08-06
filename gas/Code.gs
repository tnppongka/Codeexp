/**
 * ระบบจัดตารางเวร "ห้องยาสบปราบ"
 * Google Apps Script Backend
 * =====================================
 */

// ============ GLOBAL CONFIG ============
const SPREADSHEET_ID = ''; // ใส่ ID ของ Google Sheet ที่นี่
const LINE_CHANNEL_TOKEN = ''; // ใส่ LINE Channel Access Token

// ============ SHEET HELPERS ============
function getSheet(name) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(name);
  if (!sheet) {
    sheet = ss.insertSheet(name);
  }
  return sheet;
}

function getAllData(sheetName) {
  const sheet = getSheet(sheetName);
  const data = sheet.getDataRange().getValues();
  if (data.length <= 1) return [];
  const headers = data[0];
  return data.slice(1).map(row => {
    const obj = {};
    headers.forEach((h, i) => { obj[h] = row[i]; });
    return obj;
  });
}

function addRow(sheetName, rowData) {
  const sheet = getSheet(sheetName);
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const row = headers.map(h => rowData[h] || '');
  sheet.appendRow(row);
  return true;
}

function updateRow(sheetName, matchField, matchValue, updateData) {
  const sheet = getSheet(sheetName);
  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  const matchIdx = headers.indexOf(matchField);
  
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][matchIdx]) === String(matchValue)) {
      const updateIdxs = Object.keys(updateData).map(k => headers.indexOf(k));
      updateIdxs.forEach((colIdx, idx) => {
        const keys = Object.keys(updateData);
        sheet.getRange(i + 1, colIdx + 1).setValue(updateData[keys[idx]]);
      });
      return true;
    }
  }
  return false;
}

function deleteRow(sheetName, matchField, matchValue) {
  const sheet = getSheet(sheetName);
  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  const matchIdx = headers.indexOf(matchField);
  
  for (let i = data.length - 1; i >= 1; i--) {
    if (String(data[i][matchIdx]) === String(matchValue)) {
      sheet.deleteRow(i + 1);
      return true;
    }
  }
  return false;
}

function generateId(prefix) {
  return prefix + '_' + Utilities.getUuid().substring(0, 8) + '_' + Date.now();
}

// ============ INITIALIZE SHEETS ============
function initializeSheets() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // 1. users sheet
  let usersSheet = ss.getSheetByName('users');
  if (!usersSheet) {
    usersSheet = ss.insertSheet('users');
    usersSheet.getRange(1, 1, 1, 7).setValues([['user_id', 'fullname', 'position', 'rate_per_hour', 'line_user_id', 'is_admin', 'password']]);
    // Add demo admin
    usersSheet.appendRow(['USR_001', 'Admin', 'ผู้ดูแลระบบ', 0, '', true, 'admin1234']);
  }
  
  // 2. duty_shifts sheet
  let shiftsSheet = ss.getSheetByName('duty_shifts');
  if (!shiftsSheet) {
    shiftsSheet = ss.insertSheet('duty_shifts');
    shiftsSheet.getRange(1, 1, 1, 8).setValues([['shift_id', 'shift_name', 'shift_type', 'hours', 'time_start', 'time_end', 'rate_multiplier', 'notify_enabled']]);
    // Default shifts
    shiftsSheet.appendRow(['SHIFT_NORMAL', 'เวรวันทำงานปกติ', 'วันปกติ', 4, '16:00', '20:00', 1, true]);
    shiftsSheet.appendRow(['SHIFT_WEEKEND', 'เวรเสาร์-อาทิตย์/นักขัตฤกษ์', 'วันหยุด', 8, '08:00', '16:00', 1.5, true]);
    shiftsSheet.appendRow(['SHIFT_DOC_NORMAL', 'เวรเอกสาร (วันปกติ)', 'เวรเอกสาร', 4, '16:00', '20:00', 1, true]);
    shiftsSheet.appendRow(['SHIFT_DOC_WEEKEND', 'เวรเอกสาร (วันหยุด)', 'เวรเอกสาร', 8, '08:30', '16:30', 1.5, true]);
  }
  
  // 3. bookings sheet
  let bookingsSheet = ss.getSheetByName('bookings');
  if (!bookingsSheet) {
    bookingsSheet = ss.insertSheet('bookings');
    bookingsSheet.getRange(1, 1, 1, 8).setValues([['booking_id', 'duty_date', 'shift_type', 'user_id', 'booked_by', 'booking_time', 'status', 'notes']]);
  }
  
  // 4. notifications sheet
  let notifSheet = ss.getSheetByName('notifications');
  if (!notifSheet) {
    notifSheet = ss.insertSheet('notifications');
    notifSheet.getRange(1, 1, 1, 4).setValues([['notif_id', 'user_id', 'shift_type', 'is_active']]);
  }
  
  // 5. holidays sheet
  let holidaysSheet = ss.getSheetByName('holidays');
  if (!holidaysSheet) {
    holidaysSheet = ss.insertSheet('holidays');
    holidaysSheet.getRange(1, 1, 1, 4).setValues([['holiday_date', 'holiday_name', 'holiday_type', 'year']]);
    // Add 2568 (2025) Thai holidays
    const holidays2568 = [
      ['2025-01-01', 'วันขึ้นปีใหม่', 'national_holiday', 2568],
      ['2025-02-12', 'วันมาฆบูชา', 'national_holiday', 2568],
      ['2025-04-06', 'วันจักรี', 'national_holiday', 2568],
      ['2025-04-13', 'วันสงกรานต์', 'national_holiday', 2568],
      ['2025-04-14', 'วันสงกรานต์', 'national_holiday', 2568],
      ['2025-04-15', 'วันสงกรานต์', 'national_holiday', 2568],
      ['2025-05-01', 'วันแรงงาน', 'national_holiday', 2568],
      ['2025-05-05', 'วันฉัตรมงคล', 'national_holiday', 2568],
      ['2025-05-12', 'วันวิสาขบูชา', 'national_holiday', 2568],
      ['2025-05-30', 'วันอาสาฬหบูชา', 'national_holiday', 2568],
      ['2025-07-11', 'วันเข้าพรรษา', 'national_holiday', 2568],
      ['2025-07-28', 'วันเฉลิมพระชนมพรรษา ร.10', 'national_holiday', 2568],
      ['2025-08-12', 'วันแม่แห่งชาติ', 'national_holiday', 2568],
      ['2025-10-13', 'วันคล้ายวันสวรรคต ร.9', 'national_holiday', 2568],
      ['2025-10-23', 'วันปิยมหาราช', 'national_holiday', 2568],
      ['2025-12-05', 'วันพ่อแห่งชาติ', 'national_holiday', 2568],
      ['2025-12-10', 'วันรัฐธรรมนูญ', 'national_holiday', 2568],
      ['2025-12-31', 'วันสิ้นปี', 'national_holiday', 2568],
    ];
    holidays2568.forEach(h => holidaysSheet.appendRow(h));
  }
  
  // 6. egp sheet
  let egpSheet = ss.getSheetByName('egp');
  if (!egpSheet) {
    egpSheet = ss.insertSheet('egp');
    egpSheet.getRange(1, 1, 1, 1).setValues([['project_no']]);
  }
  
  // 7. settings sheet
  let settingsSheet = ss.getSheetByName('settings');
  if (!settingsSheet) {
    settingsSheet = ss.insertSheet('settings');
    settingsSheet.getRange(1, 1, 1, 2).setValues([['key', 'value']]);
    settingsSheet.appendRow(['buddhist_offset', '543']);
    settingsSheet.appendRow(['line_channel_access_token', '']);
    settingsSheet.appendRow(['line_notify_group_id', '']);
    settingsSheet.appendRow(['organization_name', 'ห้องยาสบปราบ']);
  }
  
  return 'สร้างชีททั้งหมดเรียบร้อยแล้ว';
}

// ============ API: USERS ============
function apiGetUsers() {
  return getAllData('users');
}

function apiAddUser(data) {
  data.user_id = generateId('USR');
  return addRow('users', data);
}

function apiUpdateUser(userId, data) {
  return updateRow('users', 'user_id', userId, data);
}

function apiDeleteUser(userId) {
  return deleteRow('users', 'user_id', userId);
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

// ============ API: DUTY SHIFTS ============
function apiGetShifts() {
  return getAllData('duty_shifts');
}

function apiUpdateShift(shiftId, data) {
  return updateRow('duty_shifts', 'shift_id', shiftId, data);
}

// ============ API: BOOKINGS ============
function apiGetBookings(year, month) {
  const allBookings = getAllData('bookings');
  if (year && month) {
    const yearNum = parseInt(year);
    const monthNum = parseInt(month);
    const christianYear = yearNum - 543;
    const prefix = christianYear + '-' + String(monthNum).padStart(2, '0');
    return allBookings.filter(b => {
      const dateStr = b.duty_date;
      if (typeof dateStr === 'string') return dateStr.startsWith(prefix);
      if (dateStr instanceof Date) {
        const d = new Date(dateStr);
        return d.getFullYear() === christianYear && d.getMonth() === monthNum - 1;
      }
      return false;
    });
  }
  return allBookings;
}

function apiCheckBooking(date, shiftType) {
  const dateStr = formatDateKey(date);
  const allBookings = getAllData('bookings');
  const existing = allBookings.find(b => {
    const bDate = b.duty_date instanceof Date 
      ? formatDateKey(b.duty_date) 
      : String(b.duty_date).substring(0, 10);
    return bDate === dateStr && b.status !== 'cancelled';
  });
  if (existing) {
    return { booked: true, bookedBy: existing.booked_by, booking: existing };
  }
  return { booked: false };
}

function apiBookDuty(bookingData) {
  // Check if already booked
  const check = apiCheckBooking(bookingData.duty_date, bookingData.shift_type);
  if (check.booked) {
    return { success: false, message: 'วันที่ ' + bookingData.duty_date + ' มีผู้จองแล้ว โดย ' + check.bookedBy };
  }
  
  bookingData.booking_id = generateId('BK');
  bookingData.booking_time = new Date().toISOString();
  bookingData.status = 'confirmed';
  
  addRow('bookings', bookingData);
  
  // Send LINE notification
  try {
    sendLineNotification(bookingData);
  } catch(e) {
    Logger.log('LINE notification error: ' + e.message);
  }
  
  return { success: true, bookingId: bookingData.booking_id };
}

function apiUpdateBooking(bookingId, data) {
  return updateRow('bookings', 'booking_id', bookingId, data);
}

function apiDeleteBooking(bookingId) {
  return deleteRow('bookings', 'booking_id', bookingId);
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
  return addRow('holidays', data);
}

function apiDeleteHoliday(date) {
  return deleteRow('holidays', 'holiday_date', date);
}

function apiGetMonthHolidays(year, month) {
  const christianYear = parseInt(year) - 543;
  const allHolidays = getAllData('holidays');
  const prefix = christianYear + '-' + String(parseInt(month)).padStart(2, '0');
  
  // Weekend holidays
  const daysInMonth = new Date(christianYear, parseInt(month), 0).getDate();
  const holidays = [];
  
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
  
  // National holidays
  allHolidays.forEach(h => {
    const hDate = h.holiday_date instanceof Date ? formatDateKey(h.holiday_date) : String(h.holiday_date).substring(0, 10);
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
  // Check if exists
  const existing = getAllData('notifications');
  const match = existing.find(n => n.user_id === data.user_id && n.shift_type === data.shift_type);
  if (match) {
    return updateRow('notifications', 'notif_id', match.notif_id, { is_active: data.is_active });
  }
  data.notif_id = generateId('NOT');
  return addRow('notifications', data);
}

// ============ API: EGP ============
function apiGetEgpProjects() {
  return getAllData('egp');
}

function apiAddEgpProject(data) {
  return addRow('egp', data);
}

// ============ API: SETTINGS ============
function apiGetSettings() {
  return getAllData('settings');
}

function apiUpdateSetting(key, value) {
  return updateRow('settings', 'key', key, { value: value });
}

// ============ API: ANALYSIS ============
function apiGetDocDutyUsers(year, month) {
  const bookings = apiGetBookings(year, month);
  const docBookings = bookings.filter(b => b.shift_type && b.shift_type.indexOf('เวรเอกสาร') !== -1);
  
  // Get unique users
  const users = [...new Map(docBookings.map(b => [b.user_id, b])).values()];
  return users.map(u => ({
    user_id: u.user_id,
    booked_by: u.booked_by,
    dates: docBookings.filter(b => b.user_id === u.user_id).map(b => b.duty_date).sort()
  }));
}

function apiAnalyzeDocDuty(userId, year, month) {
  const bookings = apiGetBookings(year, month);
  const shifts = getAllData('duty_shifts');
  
  // Get user's doc duty dates
  const userBookings = bookings.filter(b => 
    b.user_id === userId && b.shift_type && b.shift_type.indexOf('เวรเอกสาร') !== -1
  );
  
  if (userBookings.length === 0) {
    return { dates: [] };
  }
  
  // Get EGP projects
  const projects = getAllData('egp').map(p => p.project_no).filter(Boolean);
  const totalProjects = projects.length;
  const projectsPerDay = Math.ceil(totalProjects / userBookings.length);
  
  // Analyze each date
  const result = userBookings.map((booking, idx) => {
    const date = booking.duty_date instanceof Date ? booking.duty_date : new Date(booking.duty_date);
    const dayOfWeek = date.getDay();
    const dateStr = date instanceof Date ? formatDateKey(date) : String(booking.duty_date).substring(0, 10);
    
    // Determine shift type for this date
    const isWeekend = dayOfWeek === 0 || dayOfWeek === 6;
    const shiftName = isWeekend ? 'SHIFT_DOC_WEEKEND' : 'SHIFT_DOC_NORMAL';
    const shiftInfo = shifts.find(s => s.shift_id === shiftName);
    
    // Distribute projects
    const startIdx = idx * projectsPerDay;
    const dayProjects = projects.slice(startIdx, startIdx + projectsPerDay);
    
    return {
      date: dateStr,
      dateThai: formatDateThai(date),
      shift_name: shiftInfo ? shiftInfo.shift_name : shiftName,
      time_start: shiftInfo ? shiftInfo.time_start : '',
      time_end: shiftInfo ? shiftInfo.time_end : '',
      hours: shiftInfo ? shiftInfo.hours : 0,
      projects: dayProjects,
      project_count: dayProjects.length
    };
  });
  
  return { dates: result, total_projects: totalProjects, projects_per_day: projectsPerDay };
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
      const shift = shifts.find(s => s.shift_id === b.shift_type);
      return {
        ...b,
        shift_name: shift ? shift.shift_name : b.shift_type,
        time_start: shift ? shift.time_start : '',
        time_end: shift ? shift.time_end : '',
        hours: shift ? shift.hours : 0
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
    const userId = b.user_id;
    if (!userSummary[userId]) {
      userSummary[userId] = {
        user_id: userId,
        booked_by: b.booked_by,
        days: 0,
        total_hours: 0,
        total_payment: 0,
        shifts: {}
      };
    }
    
    const shift = shifts.find(s => s.shift_id === b.shift_type);
    const hours = shift ? shift.hours : 0;
    const multiplier = shift ? shift.rate_multiplier : 1;
    
    userSummary[userId].days += 1;
    userSummary[userId].total_hours += hours;
    
    const user = users.find(u => u.user_id === userId);
    const ratePerHour = user ? parseFloat(user.rate_per_hour) || 0 : 0;
    const payment = hours * ratePerHour * multiplier;
    userSummary[userId].total_payment += payment;
    
    if (!userSummary[userId].shifts[b.shift_type]) {
      userSummary[userId].shifts[b.shift_type] = { count: 0, hours: 0, payment: 0 };
    }
    userSummary[userId].shifts[b.shift_type].count += 1;
    userSummary[userId].shifts[b.shift_type].hours += hours;
    userSummary[userId].shifts[b.shift_type].payment += payment;
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
  
  const myBookings = bookings.filter(b => b.user_id === userId);
  
  const normalDays = myBookings.filter(b => b.shift_type === 'SHIFT_NORMAL').map(b => b.duty_date);
  const weekendDays = myBookings.filter(b => b.shift_type === 'SHIFT_WEEKEND').map(b => b.duty_date);
  const docNormalDays = myBookings.filter(b => b.shift_type === 'SHIFT_DOC_NORMAL').map(b => b.duty_date);
  const docWeekendDays = myBookings.filter(b => b.shift_type === 'SHIFT_DOC_WEEKEND').map(b => b.duty_date);
  
  return {
    user_id: userId,
    bookings: myBookings.map(b => {
      const shift = shifts.find(s => s.shift_id === b.shift_type);
      return {
        ...b,
        shift_name: shift ? shift.shift_name : b.shift_type,
        hours: shift ? shift.hours : 0
      };
    }),
    normal_days: normalDays,
    weekend_days: weekendDays,
    doc_normal_days: docNormalDays,
    doc_weekend_days: docWeekendDays,
    total_days: myBookings.length,
    total_hours: myBookings.reduce((sum, b) => {
      const shift = shifts.find(s => s.shift_id === b.shift_type);
      return sum + (shift ? shift.hours : 0);
    }, 0)
  };
}

// ============ CALENDAR HELPER ============
function apiGetCalendar(year, month) {
  const christianYear = parseInt(year) - 543;
  const monthNum = parseInt(month);
  const firstDay = new Date(christianYear, monthNum - 1, 1);
  const lastDay = new Date(christianYear, monthNum, 0);
  const daysInMonth = lastDay.getDate();
  const startDayOfWeek = firstDay.getDay(); // 0=Sun
  
  // Get holidays for this month
  const holidays = apiGetMonthHolidays(year, month);
  const holidayMap = {};
  holidays.forEach(h => { holidayMap[h.date] = h; });
  
  // Get bookings for this month
  const bookings = apiGetBookings(year, month);
  const bookingMap = {};
  bookings.forEach(b => {
    const dateStr = b.duty_date instanceof Date ? formatDateKey(b.duty_date) : String(b.duty_date).substring(0, 10);
    bookingMap[dateStr] = b;
  });
  
  // Build calendar days
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
    
    const dayInfo = {
      day: d,
      date: dateStr,
      dayOfWeek: dayOfWeek,
      isWeekend: isWeekend,
      isHoliday: !!holiday,
      holidayName: holiday ? holiday.name : '',
      booking: booking || null,
      isBooked: !!booking
    };
    
    days.push(dayInfo);
    
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

// ============ LINE NOTIFICATION ============
function sendLineNotification(bookingData) {
  const settings = getAllData('settings');
  const tokenSetting = settings.find(s => s.key === 'line_channel_access_token');
  const token = tokenSetting ? tokenSetting.value : '';
  const groupId = settings.find(s => s.key === 'line_notify_group_id');
  const targetId = groupId ? groupId.value : '';
  
  if (!token) return;
  
  const shifts = getAllData('duty_shifts');
  const shift = shifts.find(s => s.shift_id === bookingData.shift_type);
  const shiftName = shift ? shift.shift_name : bookingData.shift_type;
  
  const dateObj = bookingData.duty_date instanceof Date ? bookingData.duty_date : new Date(bookingData.duty_date);
  const dateThai = formatDateThai(dateObj);
  
  const message = `📋 แจ้งเตือนการจองเวร\n\n` +
    `👤 ผู้จอง: ${bookingData.booked_by}\n` +
    `📅 วันที่: ${dateThai}\n` +
    `🔄 ประเภทเวร: ${shiftName}\n` +
    `⏰ เวลา: ${shift ? shift.time_start + '-' + shift.time_end : ''}`;
  
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
      })
    };
    UrlFetchApp.fetch(url, options);
  } catch(e) {
    Logger.log('LINE API Error: ' + e.message);
  }
}

// ============ HELPER FUNCTIONS ============
function formatDateKey(date) {
  const y = date.getFullYear();
  const m = String(date.getMonth() + 1).padStart(2, '0');
  const d = String(date.getDate()).padStart(2, '0');
  return y + '-' + m + '-' + d;
}

function formatDateThai(date) {
  const thaiMonths = ['ม.ค.', 'ก.พ.', 'มี.ค.', 'เม.ย.', 'พ.ค.', 'มิ.ย.', 
                      'ก.ค.', 'ส.ค.', 'ก.ย.', 'ต.ค.', 'พ.ย.', 'ธ.ค.'];
  const d = date.getDate();
  const m = thaiMonths[date.getMonth()];
  const y = date.getFullYear() + 543;
  return d + ' ' + m + ' ' + y;
}

function getThaiMonthName(monthNum) {
  const months = ['', 'มกราคม', 'กุมภาพันธ์', 'มีนาคม', 'เมษายน', 'พฤษภาคม', 'มิถุนายน',
                  'กรกฎาคม', 'สิงหาคม', 'กันยายน', 'ตุลาคม', 'พฤศจิกายน', 'ธันวาคม'];
  return months[monthNum] || '';
}

// ============ WEB APP SERVING ============
function doGet(e) {
  return HtmlService.createTemplateFromFile('index')
    .evaluate()
    .setTitle('ระบบจัดตารางเวร - ห้องยาสบปราบ')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function doPost(e) {
  try {
    const params = JSON.parse(e.postData.contents);
    const action = params.action;
    const data = params.data || {};
    
    let result;
    switch(action) {
      // Auth
      case 'login': result = apiLogin(data.fullname, data.password); break;
      
      // Users
      case 'getUsers': result = apiGetUsers(); break;
      case 'addUser': result = apiAddUser(data); break;
      case 'updateUser': result = apiUpdateUser(data.user_id, data); break;
      case 'deleteUser': result = apiDeleteUser(data.user_id); break;
      
      // Shifts
      case 'getShifts': result = apiGetShifts(); break;
      case 'updateShift': result = apiUpdateShift(data.shift_id, data); break;
      
      // Bookings
      case 'getBookings': result = apiGetBookings(data.year, data.month); break;
      case 'checkBooking': result = apiCheckBooking(data.duty_date, data.shift_type); break;
      case 'bookDuty': result = apiBookDuty(data); break;
      case 'updateBooking': result = apiUpdateBooking(data.booking_id, data); break;
      case 'deleteBooking': result = apiDeleteBooking(data.booking_id); break;
      
      // Calendar
      case 'getCalendar': result = apiGetCalendar(data.year, data.month); break;
      
      // Holidays
      case 'getHolidays': result = apiGetHolidays(data.year); break;
      case 'getMonthHolidays': result = apiGetMonthHolidays(data.year, data.month); break;
      case 'addHoliday': result = apiAddHoliday(data); break;
      case 'deleteHoliday': result = apiDeleteHoliday(data.holiday_date); break;
      
      // Notifications
      case 'getNotifications': result = apiGetNotifications(); break;
      case 'setNotification': result = apiSetNotification(data); break;
      
      // EGP
      case 'getEgpProjects': result = apiGetEgpProjects(); break;
      case 'addEgpProject': result = apiAddEgpProject(data); break;
      
      // Settings
      case 'getSettings': result = apiGetSettings(); break;
      case 'updateSetting': result = apiUpdateSetting(data.key, data.value); break;
      
      // Analysis
      case 'getDocDutyUsers': result = apiGetDocDutyUsers(data.year, data.month); break;
      case 'analyzeDocDuty': result = apiAnalyzeDocDuty(data.user_id, data.year, data.month); break;
      
      // Reports
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
