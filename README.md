# Dance Studio Admin Portal

A mobile-friendly web app for managing class check-ins, monthly passes, and attendance analytics at a dance studio. Built entirely with Google Apps Script, served as a Web App from a Google Sheet.

---

## Features

### Check-In Portal
- Search students by name with live autocomplete (pulls from a student email database)
- Select class type with live pricing loaded from the spreadsheet template
- Regular and Student/Senior pricing toggle
- Payment method selection (Venmo, Cash, Square, Other) with Square fee calculation
- Writes to the correct date tab in the master spreadsheet without touching formula columns

### Monthly Pass Management
- View all active monthly pass holders with classes remaining
- One-tap check-in that decrements the pass count and logs the date
- Passes with zero classes remaining are automatically hidden
- Monthly passes tracked separately from drop-in attendance

### Attendance Analytics
- Tonight's class breakdown (Total, Level 0, Level 1, Level 2, Social)
- Year-to-date totals across all date tabs
- Top 10 Attendees leaderboard (excludes Social Dance Only and placeholder rows)
- Top 10 Social Dancers leaderboard (Social Dance Only purchases only)

---

## Tech Stack

| Layer | Technology |
|---|---|
| Backend | Google Apps Script (server-side JS) |
| Frontend | Bootstrap 5 + Bootstrap Icons, vanilla JS |
| Data store | Google Sheets (one tab per class date) |
| Auth | Google account (script runs as owner) |
| Hosting | Google Apps Script Web App deployment |

---

## Project Structure

```
Code.gs          — Backend: check-in logic, pricing, monthly passes, student DB
Attendance.gs    — Backend: attendance analytics and top-10 leaderboards
Index.html       — Frontend: full single-page app (Bootstrap 5, vanilla JS)
Config.gs        — NOT committed — holds your Spreadsheet ID and Logo file ID
                   (copy Config.gs.example and fill in your values)
```

---

## Setup

### 1. Copy the spreadsheet template
Create a Google Sheet with the following tab structure:
- **Date tabs** — named `M/D/YYYY` (e.g. `5/15/2026`); check-ins start at row 3
- **Weekly Template tab** — columns Q & R hold class names and prices (rows 6–25) and payment methods (rows 34–37)
- **Monthlies tab** — auto-created by the app on first use
- **StudentEmails tab** — auto-created by the app on first use

### 2. Create the Apps Script project
1. Open your Google Sheet → **Extensions → Apps Script**
2. Create three files: `Code.gs`, `Attendance.gs`, `Index.html`
3. Paste the contents of each file from this repo

### 3. Configure credentials
1. Copy `Config.gs.example` → `Config.gs` inside the Apps Script editor
2. Fill in your `SPREADSHEET_ID` (from the sheet URL) and `LOGO_FILE_ID` (Google Drive file ID for your logo)
3. `Config.gs` is listed in `.gitignore` and should never be committed

### 4. Deploy
1. **Deploy → New deployment → Web App**
2. Set "Execute as" to **Me** and "Who has access" to your organization or specific users
3. Copy the deployment URL and share with your team

### 5. Verify
Run `testSheetConnection()` from the Apps Script editor to confirm the sheet is reachable and pricing options load correctly.

---

## Key Design Decisions

**Single round-trip config load** — `getConfig()` returns pricing and payment methods in one call on page load, avoiding repeated backend calls as the admin fills out the form.

**Live pricing from the sheet** — Class options and prices are read directly from the weekly template tab, so the app automatically reflects any pricing changes without a code update.

**Validation snapshot/restore** — Column B and D use "Reject input" data validation. The app snapshots each cell's validation rule, clears it, writes the value, then restores the original rule — preserving dropdown colors and styling.

**Skip row filtering** — Rows where column A contains "skip" are excluded from all attendance counts and leaderboards, preventing template placeholder rows from inflating stats.

**Separate social leaderboard** — "Social Dance Only" attendees are counted in totals but shown in their own Top 10 list, keeping the main leaderboard focused on class students.

---

## Screenshots

> _Add screenshots of the check-in form, monthly pass list, and attendance analytics page here._

---

## License

MIT
