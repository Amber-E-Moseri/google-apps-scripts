# Drive File Alert

Automatically notifies the BLWCANADA media head by email whenever a file is uploaded to Google Drive — eliminating the manual habit of checking folders and following up with contributors to confirm uploads.

No third-party apps. No servers. Runs entirely inside your Google account for free.

---

## What it does

- Scans all folders in your Google Drive on a schedule
- Sends an email when a new file appears with the file name, folder, uploader, and a direct link
- Batches multiple uploads from the same person within **7 minutes** into one notification
- Smart schedule — checks more frequently when it matters:

| Day | Window | Frequency |
|-----|--------|-----------|
| Sunday | 12pm – 6pm | Every 15 minutes |
| Monday – Saturday | All day | Every 2 hours |

---

## Example notifications

**Single file:**
```
Subject: Drive: 1 new file from john@company.com

john@company.com added "Q3 Budget.xlsx" to Finance.
https://drive.google.com/file/d/abc123
```

**Multiple files from same person within 7 minutes:**
```
Subject: Drive: 4 new files from jane@company.com

jane@company.com added 4 files to Finance, HR, Legal.
```

---

## Setup

### 1. Open Google Apps Script
Go to [script.google.com](https://script.google.com) and click **New project**.

### 2. Paste the script
Delete the default content and paste the contents of `DriveFileAlert.gs`. Update the email address at the top:

```js
const ALERT_EMAIL = 'your@email.com';
```

Save with `Ctrl+S`.

### 3. Run once to grant permissions
Select `checkOtherDays` from the function dropdown and click **▶ Run**. Follow the permissions dialog.

### 4. Create two triggers

Go to the **Triggers** page (clock icon in the sidebar) and add:

| Trigger | Function | Type | Interval |
|---------|----------|------|----------|
| 1 | `checkSunday` | Time-driven | Every 15 minutes |
| 2 | `checkOtherDays` | Time-driven | Every 2 hours |

### 5. Test it
Upload a file to any Drive folder, then manually run `checkOtherDays`. You should receive an email within seconds.

---

## How it works

Uses `PropertiesService` to persist state between runs — remembers every file ID it has seen so it never sends duplicate alerts. New files sit in a pending queue for 7 minutes before sending, allowing uploads from the same person to batch into one notification.

---

## Quota usage (Google free tier)

| Resource | Limit | This script |
|----------|-------|-------------|
| Triggers | 20 | 2 |
| Daily runtime | 6 hours | ~2 seconds per run |
| Emails per day | 100 | Only sent on changes |

---

## Files

```
DriveFileAlert.gs   — the Apps Script source
README.md           — this file
```

---

## License

MIT
