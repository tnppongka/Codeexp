# Cross-reference analysis

## All frontend calls have matching backend functions
- No missing functions in backend for frontend calls

## Backend functions not used by frontend (acceptable - utility/reserved):
- apiAddEgpProject - reserved for future EGP management
- apiDeleteEgpProject - reserved
- apiGetBookings - not directly called but data is accessed via other APIs
- apiGetEgpProjects - reserved
- apiGetMonthHolidays - used internally by apiGetCalendar
- apiUpdateBooking - reserved for future admin edit

## Issues found:
1. In the doPost handler, the switch case uses `action` field but frontend uses `google.script.run.apiXxx()` directly
   - GAS Web Apps with HtmlService use google.script.run, NOT doPost
   - The doPost handler is only for external API calls, not internal SPA navigation
   - This is CORRECT - frontend uses google.script.run directly, doPost is for external integration

## Architecture confirmed:
- Frontend calls backend functions directly via google.script.run.apiXxx()
- doPost is a bonus for external API integration
- No mismatch issues found
