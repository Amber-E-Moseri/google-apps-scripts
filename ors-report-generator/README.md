# BLW Combined Report Suite

A full Google Apps Script reporting suite for BLW Canada follow-up and first-timer data. Reduced a 2.5-hour weekly manual reporting process to under 5 minutes.

**Before:** Manually compiling notes, membership counts, and first-timer data from multiple Elvanto exports into formatted subgroup reports — every week.
**After:** One click generates all 6 subgroup reports simultaneously, with automatic grading, colour-coded health bands, PDF export, and email delivery built in.

---

## What it does

### 📦 Comprehensive Reports
- **Generate All Reports** — one button generates all 6 subgroup reports in a single run, prompting for label, grade toggle, and colour preferences
- **Email Reports as PDF** — exports any report sheet to PDF and emails it to a specified recipient, all from within Sheets

### 📊 Follow-Up Reports
Reads follow-up note data from a cleaned Elvanto export and produces:
- **Regional Overview** — subgroup-level summary table with notes, unique people reached, member counts, groups active, % reached, and optional grade
- **Subgroup Reports** — per-subgroup sheets combining follow-up stats, fellowship-level breakdown, leaders participating, and first-timer data
- **Leader Note Counter** — ranked leaderboard of notes logged per leader, grouped by subgroup with subtotals
- **Delete Report Sheets** — selectively or bulk-delete generated sheets while keeping source data protected

### 📋 First-Timer Tools
Reads CSV exports from Elvanto (Cell FTs and Service FTs) and produces:
- **FT Regional Report** — regional overview showing cell, service, and combined first-timer totals by subgroup and fellowship
- **FT Subgroup Reports** — first-timer data embedded in each subgroup's combined report
- **Diagnose FT Data** — shows how the demographics parser is resolving each row's subgroup, and flags any unmatched values
- **Setup FT Import Sheets** — creates the two import tabs with correct headers ready for CSV paste

---

## Grading system

The grade (A+ through F) is calculated from two factors:
- **% Reached** (80% weight) — unique people followed up ÷ adjusted members (total members − 1 per group to exclude the leader)
- **Group Mobilisation** (20% weight) — groups with at least one note ÷ total groups in subgroup

Health-band colours (green / yellow / orange / red) are applied to grade cells and optionally to % Reached cells.

---

## Data sources

| Sheet | Contents |
|-------|----------|
| `ELVANTO_NOTES_CLEAN` | Cleaned follow-up notes export from Elvanto |
| `GROUP` | Group directory with name, category (subgroup), and member count |
| `Cell FTs Import` | CSV export of cell first-timers from Elvanto |
| `Service FTs Import` | CSV export of service first-timers from Elvanto |

Required columns for FT import sheets: `Full Name`, `Groups`, `Demographics`, `Date Added`

---

## Setup

### 1. Open Apps Script
In your Google Sheet: **Extensions → Apps Script**

### 2. Paste the script
Delete the default content, paste `BLWReportSuite.gs`, and save (`Ctrl+S`).

### 3. Grant permissions
Run `onOpen` from the function dropdown and follow the permissions dialog.

### 4. Reload your spreadsheet
The three menus — **📦 Comprehensive Reports**, **📊 Follow-Up Reports**, and **📋 First-Timer Tools** — will appear in the top menu bar.

### 5. Set up FT import sheets (first time only)
Go to **📋 First-Timer Tools → ⚙️ Setup FT Import Sheets**. This creates the two import tabs with correct headers. Paste your Elvanto CSV exports into each tab.

---

## Subgroups

The script handles six subgroups for BLW Canada:

| Follow-Up name | First-Timer name |
|---------------|-----------------|
| BLW Canada Central-East SGA | Central East Subgroup A |
| BLW Canada Central-East SGB | Central East Subgroup B |
| BLW Canada Central SGA | Central Subgroup A |
| BLW Canada Central SGB | Central Subgroup B |
| BLW Canada West SGA | West Subgroup A |
| BLW Canada West SGB | West Subgroup B |

The crosswalk between these two naming conventions is handled automatically.

---

## Customisation

| What to change | Where |
|---------------|-------|
| Subgroup names | `CFG.SUBGROUPS` and `CFG.CATEGORY_MAP` |
| Source sheet names | `CFG.NOTES_SHEET`, `CFG.GROUP_SHEET` |
| Timezone | `CFG.TIMEZONE` |
| Grade thresholds | `gradeScore_()` and `gradeLabel_()` |
| Colour theme | `T` (follow-up) and `FT` (first-timer) objects |
| Demographics aliases | `FT_SUBGROUP_ALIASES` |

---

## Impact

| Metric | Before | After |
|--------|--------|-------|
| Time to produce all reports | ~2.5 hours | Under 5 minutes |
| Manual steps | Dozens | 1 click (Generate All) |
| Reports generated | 6 subgroup sheets | 6 subgroup + regional + leader counter |
| Distribution | Manual | PDF email delivery built in |

---

## Files

```
BLWReportSuite.gs   — the full Apps Script source (~1,200 lines)
README.md           — this file
```

---

## License

MIT
