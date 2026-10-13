# Application-Logs-Viewer

> Repository: <https://github.com/muhamaddarulhadi/Application-Logs-Viewer>

> [Website Tutorial](https://muhamaddarulhadi.github.io/Application-Logs-Viewer/)

> Created by [muhamaddarulhadi](https://github.com/muhamaddarulhadi)

**Live tutorial:** open `index.html` in a browser, or publish it with GitHub Pages (see [Project structure](#project-structure)).

`Get-AppEventReport.ps1` is a PowerShell script that reads **Error** and **Warning** events from **Event Viewer > Windows Logs > Application** for a date range you choose, then builds:

- a modern, interactive **HTML report** (charts, search, filters, detail panel, dark mode),
- a real **Excel `.xlsx`** file of the table,
- a **CSV** file of the same data.

Each event includes the date/time, level, source, event ID, category, computer, record ID, and the **General** message text (the same text you see on the General tab in Event Viewer).

---

## Requirements

| Requirement | Details |
|---|---|
| Operating system | Windows 10 / 11 or Windows Server 2016+ |
| PowerShell | Windows PowerShell 5.1 (built in) or PowerShell 7+ |
| Permissions | A normal user account is usually enough to read the Application log. Run as Administrator only if you get an access error. |
| Excel | **Not required** to create the `.xlsx`. You only need Excel (or LibreOffice) to open it. |
| Internet | **Not required.** The HTML report works fully offline. |

---

## Quick start

1. Get the project. Clone the repository (or use **Code > Download ZIP** on GitHub and extract it to the same folder):
   ```powershell
   git clone https://github.com/muhamaddarulhadi/Application-Logs-Viewer.git C:\Scripts\Application-Logs-Viewer
   ```
2. Open **PowerShell** (Start menu > type `PowerShell`).
3. Go to the folder:
   ```powershell
   cd C:\Scripts\Application-Logs-Viewer
   ```
4. If Windows blocks the script ("running scripts is disabled"), allow it for this window only:
   ```powershell
   Set-ExecutionPolicy -Scope Process -ExecutionPolicy Bypass
   ```
   If the file was downloaded, you may also need to unblock it once:
   ```powershell
   Unblock-File .\Get-AppEventReport.ps1
   ```
5. Run the script:
   ```powershell
   .\Get-AppEventReport.ps1
   ```
6. Enter the dates when asked, or press **Enter** to accept the default shown in brackets:
   ```
   Start date [default: 2026-10-06]: 2026-10-01
   End date [default: 2026-10-13]:
   ```
7. When it finishes, the HTML report opens in your browser and the file locations are printed in the console.

---

## Parameters

| Parameter | Type | Default | Description |
|---|---|---|---|
| `-StartDate` | date | Prompted (default: 7 days ago) | First day to include. Accepts `2026-10-01`, `10/01/2026`, or a date with time like `"2026-10-01 08:00"`. |
| `-EndDate` | date | Prompted (default: today) | Last day to include. A date without a time covers the **whole day** (up to 23:59:59). Must not be earlier than `-StartDate`. |
| `-OutputFolder` | path | `Desktop\EventReports` | Where the HTML, CSV and XLSX files are saved. The folder is created if it does not exist. |
| `-IncludeCritical` | switch | Off | Also include **Critical** events. By default only Error and Warning are collected. |
| `-OpenReport` | switch | Off | Open the HTML report in your browser when done. The report also opens automatically if you were prompted for dates (no `-StartDate` given). |

> If you leave out `-StartDate` or `-EndDate`, the script asks for them interactively. If you pass both, it runs without any prompt, which is useful for scheduled tasks.

### Examples

**Interactive (prompts for the date range):**
```powershell
.\Get-AppEventReport.ps1
```

**Specific date range:**
```powershell
.\Get-AppEventReport.ps1 -StartDate 2026-10-01 -EndDate 2026-10-13
```

**Just one day:**
```powershell
.\Get-AppEventReport.ps1 -StartDate 2026-10-13 -EndDate 2026-10-13
```

**A time window within a day:**
```powershell
.\Get-AppEventReport.ps1 -StartDate "2026-10-13 14:00" -EndDate "2026-10-13 15:00"
```

**Include Critical events too:**
```powershell
.\Get-AppEventReport.ps1 -StartDate 2026-10-01 -EndDate 2026-10-13 -IncludeCritical
```

**Save to a custom folder and open the report:**
```powershell
.\Get-AppEventReport.ps1 -StartDate 2026-10-01 -EndDate 2026-10-13 -OutputFolder D:\Reports -OpenReport
```

**Last 30 days, no prompts:**
```powershell
.\Get-AppEventReport.ps1 -StartDate (Get-Date).AddDays(-30) -EndDate (Get-Date) -OpenReport
```

---

## Output files

Files are saved to `-OutputFolder` (default `Desktop\EventReports`) with a timestamp so they never overwrite each other:

| File | Contents |
|---|---|
| `AppEvents_YYYYMMDD_HHMMSS.html` | The interactive report. Single self-contained file, you can email it or open it anywhere. |
| `AppEvents_YYYYMMDD_HHMMSS.xlsx` | Excel workbook with **all** events found in the date range. |
| `AppEvents_YYYYMMDD_HHMMSS.csv` | Same data as plain CSV (UTF-8). |

### Columns in the table / Excel file

| Column | Meaning |
|---|---|
| Date | When the event was logged (`yyyy-MM-dd HH:mm:ss`, local time) |
| Level | Error, Warning (or Critical with `-IncludeCritical`) |
| Source | The application or service that logged it (e.g. `VSS`, `Application Error`) |
| EventID | The event ID number |
| Category | Task category (often `None`) |
| Computer | Computer name |
| RecordId | Unique record number in the log |
| Message | The full **General** description text |

---

## Using the HTML report

- **Summary cards** show totals for events, errors, warnings and unique sources.
- **Events per day** chart shows when problems happened. Hover a bar for exact counts.
- **Severity split** and **Top sources** show the biggest contributors. Click a source to filter the table to it.
- **Search box** searches source, event ID, category and message text.
- **All / Errors / Warnings tabs** and the **source dropdown** narrow the table.
- **Click a column header** to sort. Click again to reverse.
- **Click any row** to open the detail panel with all fields and the full General message.
  - **Copy** copies the message text.
  - **Previous / Next** (or the left / right arrow keys) move between events. **Esc** closes the panel.
- **Moon / sun button** in the top corner switches dark and light mode (your choice is remembered).
- **Rows** selector sets how many rows show per page (25, 50, 100, 250).

### Downloading Excel from the report

Click **Download Excel (.xlsx)** in the toolbar.

- It exports the rows that match your **current filters**, not just the visible page. For example, filter to *Errors* and source *VSS*, then download to get only those rows.
- To export everything, clear the search box, set the tab to **All** and the source to **All sources** first.
- The file has a colored header, colored Level cells, a frozen header row, filter buttons on every column, and wrapped message text.
- The **CSV** button works the same way.

> The script also writes a full, unfiltered `.xlsx` to the output folder every time it runs, so you always have one without opening the report.

---

## Troubleshooting

| Problem | What to do |
|---|---|
| `running scripts is disabled on this system` | Run `Set-ExecutionPolicy -Scope Process -ExecutionPolicy Bypass`, then run the script again in the same window. |
| `File ... cannot be loaded. The file is not digitally signed` | Run `Unblock-File .\Get-AppEventReport.ps1`. |
| `No errors or warnings found in that date range.` | The range has no matching events. Try a wider range, or add `-IncludeCritical`. Also check the dates are in the past and `-EndDate` is not before `-StartDate`. |
| Some messages say `(Description unavailable) Raw data: ...` | The program that logged the event has no message file installed (often an uninstalled app). The raw event data is shown instead. |
| Access denied / cannot read the log | Open PowerShell with **Run as administrator** and try again. |
| Script is slow on a huge range | Windows reads the log sequentially. Use a shorter date range or run it during quiet hours. Tens of thousands of events can take a minute or more. |
| Report is slow to open in the browser | Very large reports (50,000+ events) are heavy. Use a narrower date range, or use the `.xlsx` / `.csv` instead. |
| Desktop folder is in OneDrive and the path looks odd | That is normal. Use `-OutputFolder C:\Reports` to save somewhere else. |
| Excel shows a very long message cut off | Excel limits a cell to about 32,000 characters. The script trims anything longer. Full text is still in the HTML report's detail panel. |

---

## Good to know

- **Windows only.** The script uses `Get-WinEvent`, which reads the Windows event log.
- **Only the Application log** is read. It does not touch Security, System, or other logs.
- **Read-only.** The script never changes or clears any event logs.
- **Dates use the local time** of the computer running the script.
- The `.xlsx` is built by the script itself (no Excel, no extra modules, no internet), so it works on locked-down machines.
- To run it automatically, create a Windows **Task Scheduler** task that runs:
  ```
  powershell.exe -NoProfile -ExecutionPolicy Bypass -Command "& 'C:\Scripts\Application-Logs-Viewer\Get-AppEventReport.ps1' -StartDate (Get-Date).AddDays(-1) -EndDate (Get-Date)"
  ```
  Because both dates are passed, no prompt appears and the report is not opened. Use `-Command` (not `-File`) so the `(Get-Date)` parts are evaluated.

---

## Project structure

```
Application-Logs-Viewer/
├── Get-AppEventReport.ps1     The script
├── README.md                  This guide
├── index.html                 Tutorial website (open it in any browser)
└── assets/
    ├── css/style.css          Website styles
    ├── js/main.js             Theme toggle, copy buttons, command builder
    ├── favicon.svg            Site icon
    ├── favicon.ico            Icon for older browsers
    └── apple-touch-icon.png   Icon for phones and tablets
```

### Publish the tutorial website with GitHub Pages

1. Push this folder to the `Application-Logs-Viewer` repository on GitHub.
2. Open the repository **Settings > Pages**.
3. Under **Build and deployment**, choose **Deploy from a branch**, select `main` and `/ (root)`, then save.
4. After a minute the site is available at `https://muhamaddarulhadi.github.io/Application-Logs-Viewer/`.

The website uses Google Fonts when online and falls back to system fonts when offline.

---

## Author

Made by [muhamaddarulhadi](https://github.com/muhamaddarulhadi) on GitHub.
