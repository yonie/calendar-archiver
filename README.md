# Calendar Archiver

A simple tool to download and archive Outlook/Office365 calendar data for record-keeping and tax purposes.

## Overview

This tool downloads ICS calendar files from online URLs (particularly Outlook/Office365) and creates organized, readable archives of calendar events filtered by month and year. Perfect for maintaining records of meetings, appointments, and business activities.

## Features

- **Online ICS Download**: Fetches calendar data directly from web URLs
- **Month/Year Filtering**: Archives only events from specified time periods
- **Readable Output**: Generates clean Markdown files for easy viewing
- **Organized Structure**: Creates year/month/day folder hierarchy
- **Change Detection**: Tracks updates between runs
- **Robust Downloads**: Handles large files and slow connections reliably

## Installation

### Requirements
- Python 3.6 or higher
- Required packages: `requests`, `icalendar`

### 1. Install Dependencies
```bash
pip install -r requirements.txt
```

Or install manually:
```bash
pip install requests icalendar
```

### 2. Configure Your Calendar URL
Copy the example configuration file and add your calendar URL:

```bash
cp example.env .env
```

Edit `.env` and add your calendar URL:
```bash
# Your Outlook/Office365 calendar ICS URL
ICS_URL=https://outlook.office365.com/owa/calendar/YOUR_CALENDAR_ID/calendar.ics

# Optional: Default year (if not specified, uses current year)
DEFAULT_YEAR=2025

# Optional: Default month (if not specified, archives all months)
# DEFAULT_MONTH=08
```

### Getting Your Calendar URL

**For Outlook/Office365:**
1. Open Outlook on the web
2. Go to Calendar
3. Click "Share" → "Publish calendar"
4. Choose "Can view all details"
5. Copy the ICS link and paste it in your `.env` file

## Usage

The tool supports multiple usage modes:

```bash
# Archive all months for the year (using .env configuration)
python calendar_archiver.py

# Archive specific month using .env URL
python calendar_archiver.py <year> <month>

# Archive specific month with custom URL
python calendar_archiver.py <ics_url> <year> <month>
```

### Examples

```bash
# Archive all months in 2025 (uses .env settings)
python calendar_archiver.py

# Archive August 2025 using .env URL
python calendar_archiver.py 2025 08

# Archive December 2024 with custom URL
python calendar_archiver.py "https://company.sharepoint.com/calendar.ics" 2024 12
```

## Output Structure

The tool creates an organized folder structure directly from the current directory:

```
2025/
├── .metadata/
│   └── 2025_08_events.json
├── 08/
│   ├── 04/
│   │   └── 2025-08-04_events.md
│   ├── 05/
│   │   └── 2025-08-05_events.md
│   └── ...
├── 09/
│   ├── 01/
│   │   └── 2025-09-01_events.md
│   └── ...
└── 12/
    ├── 15/
    │   └── 2025-12-15_events.md
    └── ...
```

### Event File Format

Each daily events file is a clean Markdown document:

```markdown
# Calendar Events - 2025-08-19

**Date:** Monday, August 19, 2025
**Total Events:** 10

---

## 1. Team Meeting

**Time:** 09:00 - 10:00  
**Duration:** 1:00:00  
**Location:** Conference Room A  
**Participants:** john@company.com, jane@company.com  
**Description:** Weekly team sync and project updates  
**Event ID:** `abc123...`

---

## 2. Client Call

**Time:** 14:30 - 15:30  
**Duration:** 1:00:00  
**Location:** Microsoft Teams Meeting  
**Description:** Quarterly business review with client  
**Event ID:** `def456...`

---
```

## Use Cases

### Business & Tax Records
- Maintain detailed meeting logs for tax deductions
- Track client interactions and billable time
- Document business travel and appointments
- Create audit trails for compliance

### Personal Organization
- Archive important appointments and events
- Keep historical records of activities
- Export calendar data for backup purposes

### Project Management
- Document project meetings and milestones
- Track team activities and collaboration
- Maintain records of stakeholder interactions

## Technical Details

### Robust Download Handling
- Progressive timeout strategy (30s, 60s, 120s)
- Automatic retry on server errors
- Optimized headers for Outlook/Office365 compatibility
- Chunked download for large calendar files
- Detailed progress reporting

### Change Detection
- MD5 hashing to detect event modifications
- Incremental updates preserve existing data
- Metadata tracking for efficient processing

### Data Processing
- Parses standard ICS/iCal format
- Extracts all relevant event details
- Handles timezone information
- Processes recurring events properly

## Troubleshooting

### Common Issues

**"Permission denied" or "Access denied"**
- Ensure your calendar is published/shared with the correct permissions
- Check that the ICS URL is publicly accessible

**"No events found"**
- Verify the year/month parameters are correct
- Check that events exist in the specified time period
- Ensure the calendar contains events (not just free/busy info)

**Download timeouts**
- The tool automatically retries with longer timeouts
- Large calendars (1000+ events) may take 1-2 minutes to download
- Check your internet connection

### Getting Help

If you encounter issues:
1. Check that your ICS URL works in a web browser
2. Verify the calendar contains events for your target month
3. Ensure you have write permissions in the current directory

## File Management

### Cleanup
- Archive old calendar folders as needed
- The `.metadata` folder tracks changes - don't delete it
- Generated calendar files are excluded from git by default

### Backup
- Calendar archive folders are self-contained
- Copy entire year folders (e.g., `2025/`) for backup
- Markdown files can be viewed in any text editor or browser

## License

This project is licensed under the MIT License - see the [LICENSE](LICENSE) file for details.

The MIT License is one of the most permissive open-source licenses, allowing you to:
- Use the software for any purpose (commercial or personal)
- Modify and distribute the software
- Include it in proprietary software
- No warranty or liability requirements

Feel free to fork, modify, and distribute this tool as needed!
