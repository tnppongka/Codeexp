# Development Notes

## Architecture
- Google Apps Script Web App (doGet/doPost)
- Google Sheets as database (7 sheets: users, duty_shifts, bookings, notifications, holidays, egp, settings)
- LINE Messaging API for notifications
- PWA support (manifest.json, standalone display)
- SPA pattern with sidebar navigation

## Files Created
1. gas/Code.gs - Full backend (CRUD + API functions + calendar + LINE notify)
2. gas/index.html - Full frontend (login + 4 pages + modals + responsive)
3. gas/manifest.json - PWA manifest
4. README.md - Setup guide

## Key Functions in Code.gs
- initializeSheets() - Create all 7 sheets with headers + seed data
- apiLogin(fullname, password) - Simple auth
- apiGetCalendar(year, month) - Build calendar with holidays + bookings
- apiBookDuty(data) - Book with duplicate check + LINE notify
- apiCheckBooking(date, shiftType) - Check if date already booked
- apiGetMonthlyReport(year, month) - Monthly duty report
- apiGetPaymentReport(year, month) - Payment evidence
- apiGetMyDuty(userId, year, month) - Personal duty summary
- apiAnalyzeDocDuty(userId, year, month) - Doc duty analysis with project distribution
- apiGetDocDutyUsers(year, month) - List users on doc duty
- sendLineNotification(bookingData) - Push to LINE OA
- apiGetShifts(), apiGetUsers(), apiGetHolidays(), apiGetNotifications(), apiGetSettings()
- Helper: formatDateKey(), formatDateThai(), getThaiMonthName()

## Key Features in index.html
- Login page (gradient background, card UI)
- Sidebar navigation (4 menu items + user info + logout)
- Booking page (calendar grid, month/year selector, summary stats, print)
- Settings page (5 tabs: users, shifts, notifications, holidays, system settings)
- Reports page (5 tabs: schedule, doc duty, doc attach, payment, my duty)
- Analysis page (dropdown user selector, project distribution table)
- Modal system for forms
- Toast notifications
- Responsive design (mobile hamburger menu)
- Print styles

## Thai Holidays 2568 (2025)
Pre-seeded in holidays sheet: New Year, Makha Bucha, Chakri Day, Songkran, Labor Day, Coronation, Wisakha Bucha, Asalha Bucha, Buddhist Lent, King's Birthday, Mother's Day, King Rama IX Memorial, Chulalongkorn Day, Father's Day, Constitution Day, New Year's Eve

## Default Shifts
- SHIFT_NORMAL: วันทำงานปกติ, 4hr, 16:00-20:00, 1x rate
- SHIFT_WEEKEND: เสาร์-อาทิตย์/นักขัตฤกษ์, 8hr, 08:00-16:00, 1.5x rate
- SHIFT_DOC_NORMAL: เวรเอกสารวันปกติ, 4hr, 16:00-20:00, 1x rate
- SHIFT_DOC_WEEKEND: เวรเอกสารวันหยุด, 8hr, 08:30-16:30, 1.5x rate

## Default Admin
- user_id: USR_001, fullname: Admin, password: admin1234, is_admin: true

## Remaining Work
- Push to GitHub repo tnppongka/Codeexp
- User needs to: provide LINE token, files 1-4 for report templates
- Test end-to-end flow
