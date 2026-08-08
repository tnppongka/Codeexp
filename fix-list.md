# Bugs & Gaps Found - v3.0 Fix List

## Backend (Code.gs) Issues

1. **is_admin stored as boolean but GAS converts to 'true'/'false' string in Sheets** — apiLogin compares with strict === on password but is_admin might be read as string. Need to normalize: `u.is_admin === true || u.is_admin === 'true'`
2. **apiGetCalendar holidayDays double counting** — weekendDays counted separately but holidayDays includes weekends. Should be: holidayDays = only national holidays (not weekends). Currently weekend days are counted in both.
3. **apiGetMonthHolidays** — returns weekend holidays as objects but calendar cells check `isHoliday && !isWeekend` for orange highlight. OK.
4. **sendLineNotification** uses `settings.find(s => s.key === ...)` but key is in column A — works.
5. **BUDGET_OFFSET as const in GAS** — GAS doesn't support ES6 const at global scope in all cases. Should use `var`.
6. **initializeSheets** uses `let s` — should use `var`.
7. **apiAddUser** — `data.is_admin === undefined` but GAS passes it as undefined when checkbox not checked — OK.
8. **apiBookDuty** — `bookingData.notes = bookingData.notes || ''` — OK.
9. **holiday_date column** — GAS stores dates as Date objects. `String(h.holiday_date).substring(0,10)` will fail for Date objects. Need to convert Date to YYYY-MM-DD.
10. **addRow** — when date is stored as Date object, `getAllData` returns Date object. String conversion needed everywhere.
11. **apiGetBookings** — same issue with duty_date being Date object.

## Frontend (index.html) Issues

12. **Missing EGP management UI** — No way to add/manage projects in egp sheet from settings page. Need to add tab.
13. **Calendar cell CSS** — `.calendar-cell.holiday` style uses `#fff7ed` but CSS variable `--warning` is defined. OK.
14. **`--warning-light` CSS variable** referenced but not defined — need to add.
15. **`--success-light` CSS variable** referenced but not defined — need to add.
16. **`--info-light` CSS variable** referenced but not defined — need to add.
17. **`--danger-light` CSS variable** referenced but not defined — need to add.
18. **`--border` CSS variable** referenced but not defined — need to add.
19. **`--text-lighter` CSS variable** referenced but not defined — need to add.
20. **`--shadow` CSS variable** referenced but not defined — need to add.
21. **`--shadow-sm` CSS variable** referenced but not defined — need to add.
22. **`--shadow-lg` CSS variable** referenced but not defined — need to add.
23. **`--radius` CSS variable** referenced but not defined — need to add.
24. **`--radius-sm` CSS variable** referenced but not defined — need to add.
25. **`--radius-xs` CSS variable** referenced but not defined — need to add.
26. **`--transition` CSS variable** referenced but not defined — need to add.
27. **`--accent` CSS variable** referenced but not defined — need to add.
28. **`--accent-light` CSS variable** referenced but not defined — need to add.
29. **`--bg` CSS variable** referenced but not defined — need to add.
30. **`--surface` CSS variable** referenced but not defined — need to add.
31. **`--warning` CSS variable** referenced but not defined — need to add.

Wait — looking at the CSS more carefully, these ARE defined in the `:root` block. Let me re-check...

Actually the CSS :root block defines:
- `--primary`, `--primary-light`, `--primary-dark` ✓
- `--accent`, `--accent-light` ✓
- `--bg`, `--surface` ✓
- `--text`, `--text-light`, `--text-lighter` ✓
- `--border` ✓
- `--success`, `--success-light` ✓
- `--danger`, `--danger-light` ✓
- `--warning`, `--warning-light` ✓
- `--info`, `--info-light` ✓
- `--radius`, `--radius-sm`, `--radius-xs` ✓
- `--shadow-sm`, `--shadow`, `--shadow-lg` ✓
- `--transition` ✓

All CSS variables are defined. No issues there.

## Real Issues to Fix:

### Backend:
1. **Date handling** — When reading from Sheets, dates come as Date objects. `String(date).substring(0,10)` gives "Thu Jan 01..." not "2025-01-01". Need proper Date→string conversion.
2. **is_admin type normalization** — Sheets stores boolean as 'true'/'false' string.
3. **holidayDays counting** — currently counts weekends + national holidays together. Should separate.

### Frontend:
4. **Add EGP Project Management** — Add tab in settings for managing egp projects (add/delete).
5. **formatDateThaiStr** — Uses `new Date(dateVal)` which works for Date objects and strings. OK.
6. **formatDateNumericStr** — Same. OK.

## Fix Priority:
1. Fix Date handling in Backend (critical bug)
2. Fix is_admin normalization (critical bug)
3. Fix holidayDays counting (minor)
4. Add EGP management UI (feature gap)
5. Everything else works
