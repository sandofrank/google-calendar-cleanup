# Google Calendar Cleanup Tool

🗓️ **Automatically delete old calendar events with smart filtering and comprehensive safety features.**

This Google Apps Script helps you clean up years of old calendar events in minutes instead of hours of manual deletion.

## ✨ Features

- **🛡️ Safe by Default**: Starts in dry-run mode (shows what would be deleted without actually deleting)
- **🧠 Smart Filtering**: Automatically skips birthdays, holidays, and important events
- **📅 Multi-Calendar Support**: Works across all your Google calendars
- **💾 Backup Creation**: Can create spreadsheet backups before deletion
- **📊 Detailed Reporting**: Shows exactly what was processed
- **⚙️ Highly Configurable**: Customize time periods, filters, and safety settings
- **📧 Email Reports**: Get summaries sent to your inbox

## 🚀 Quick Start (5 Minutes)

### Step 1: Set Up the Script
1. Go to [script.google.com](https://script.google.com)
2. Click **"New Project"**
3. Delete the default code
4. Copy and paste the code from `calendar-cleanup.gs` (in this repository)
5. Save the project (Ctrl+S) and give it a name

### Step 2: Test First
1. In the function dropdown, select **"quickTest"**
2. Click the **Run** button (▶️)
3. Grant permissions when Google asks (this is normal)
4. Check the execution log to confirm it works

### Step 3: Preview Your Cleanup
1. Change function to **"deleteOldCalendarEvents"**
2. Click **Run** (it starts in safe "dry run" mode)
3. Review what it would delete in the execution log

### Step 4: Configure (Optional)
Edit the `CONFIG` section at the top of the script:
```javascript
YEARS_TO_KEEP: 1,              // Keep events from last 1 year
DRY_RUN: true,                 // Change to false when ready for real deletion
SKIP_KEYWORDS: ['birthday'],   // Events containing these words will be kept
```

### Step 5: Run for Real (When Ready)
1. Change `DRY_RUN: false` in the CONFIG
2. Run **"deleteOldCalendarEvents"** again
3. Watch the progress in the execution log

## 📋 Configuration Options

| Setting | Description | Default |
|---------|-------------|---------|
| `YEARS_TO_KEEP` | Keep events from last X years | `1` |
| `DRY_RUN` | Preview mode (true) vs actual deletion (false) | `true` |
| `MAX_DELETES_PER_RUN` | Maximum events to delete in one run | `50` |
| `SKIP_KEYWORDS` | Never delete events containing these words | `['birthday', 'anniversary', 'holiday']` |
| `SKIP_EVENTS_WITH_ATTENDEES` | Keep meeting invites | `true` |
| `BACKUP_BEFORE_DELETE` | Create spreadsheet backup | `false` |
| `EMAIL_REPORT` | Send results to email | `false` |

## 🛡️ Safety Features

- **Starts in dry-run mode** - shows what would happen without doing it
- **Permission-based deletion** - can't delete events you don't own
- **Smart filtering** - automatically protects important event types
- **Batch processing** - limits how many events are processed at once
- **Detailed logging** - shows exactly what happened
- **Backup options** - can create spreadsheet copies before deletion

## 🔄 Automation (Optional)

Set up monthly automatic cleanup:

1. In Google Apps Script, click the **clock icon** (Triggers)
2. Click **"+ Add Trigger"**
3. Choose:
   - Function: `deleteOldCalendarEvents`
   - Event source: `Time-driven`
   - Type: `Month timer`
4. Save

## ❓ FAQ

**Q: Is this safe?**
A: Yes! It starts in preview mode and has multiple safety features. Always test with dry-run first.

**Q: Will it delete shared calendar events?**
A: Usually no - you typically can't delete events created by others.

**Q: What if I delete something important?**
A: Enable `BACKUP_BEFORE_DELETE: true` to create a spreadsheet backup first.

**Q: How long does it take?**
A: Depends on your calendar size. Usually seconds to a few minutes.

## 🐛 Troubleshooting

**"Permission denied" errors**: Normal for events you can't delete (created by others)

**"Rate limit" errors**: Script is paused automatically, just run again later

**No events deleted**: Check if `DRY_RUN` is still `true`, or all your events might be filtered out for safety

## 📄 License

MIT License - feel free to modify and share!

## 🤝 Contributing

Found a bug or have an improvement? Please open an issue or submit a pull request!

---

⭐ **Star this repository if it helped you clean up your calendar!**
