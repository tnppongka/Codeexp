# index.html Structure Notes (v3.0)

## Settings Page Tabs (5 tabs)
- tab-users: บุคลากร
- tab-shifts: ประเภทเวร
- tab-notif: การแจ้งเตือน
- tab-holidays: วันหยุด
- tab-settings: ตั้งค่าระบบ

## Reports Page Tabs (5 tabs)
- report-schedule: ตารางเวรประจำเดือน
- report-doc: ตารางเวรเอกสาร
- report-doc-attach: เอกสารแนบเวรเอกสาร
- report-payment: หลักฐานการจ่ายเงิน
- report-my-duty: เวรฉันวันไหน

## Key Functions in index.html
- renderSettingsPage() at line ~1223
- renderReportsPage() at line ~1587
- switchSettingsTab() at line ~1253
- switchReportTab() at line ~1625
- loadReport() at line ~1632
- renderBookingPage() - calendar view
- renderAnalysisPage() - doc duty analysis
- gasCall() - calls google.script.run with Promise wrapper

## EGP Management
- Backend has apiGetEgpProjects, apiAddEgpProject, apiDeleteEgpProject
- Frontend DOES NOT have a UI for managing EGP projects in settings
- NEED TO ADD: A new tab "เลขโครงการ (EGP)" in settings page for adding/deleting projects

## Bugs Fixed in Code.gs (v3.0):
1. All ES6 const/let/arrow functions converted to var/for loops (GAS compatibility)
2. Date objects from Sheets converted to YYYY-MM-DD strings in getAllData()
3. is_admin normalized to boolean in apiLogin()
4. holidayDays renamed to nationalHolidayDays (separate from weekendDays)
5. apiGetCalendar properly separates weekend vs national holiday counts
6. All .find() and .forEach() converted to for loops

## Remaining Tasks:
1. Add EGP project management tab in settings page
2. Verify all frontend functions work correctly with updated backend
3. Test the complete system
