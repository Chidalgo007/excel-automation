Excel Automation Program
A Python-powered automation tool that updates, refreshes, and manages Excel workbooks—hands-free.
Built for analysts, report managers, and data professionals who need reliable, scheduled, and repeatable Excel updates without the tedious manual steps.

✨ Features
Automatic Data Refresh
Opens a target Excel file and refreshes all queries and data connections to pull the latest available data.

Dynamic Date Updates
Modifies specific date cells to keep reporting periods accurate.
Pivot Table Refresh
Updates OLAP Pivot Tables across multiple worksheets to ensure consistent, up-to-date analytics.
Run Now or Schedule
Run instantly with one click.
Schedule for later once or recurring (weekly).

Persistent Scheduling
Powered by Windows Task Scheduler, ensuring tasks run even if:
The program is closed.
The computer is restarted.
Silent Background Execution
Runs Excel updates via a Python process in the background—no pop-ups, no manual work.
Built-In Tools
Information Button: In-app guide explaining functionality and usage.
Schedule Viewer: View and delete scheduled tasks with ease.

🛠️ Tech Stack
Language: Python 3.x

Key Libraries:
pywin32 – Excel COM automation
win10toast – Notifications
Other Windows automation modules
Platform: Windows (integrated with Windows Task Scheduler)

🚀 How It Works
Open & Refresh – Launches the specified workbook, updates all queries/connections, and refreshes pivot tables.
Update Dates – Modifies target date cells to reflect the correct reporting period.
Close & Save – Saves the updated file and closes Excel.
Automation Mode – Run now or set a persistent schedule (runs even if your PC is restarted).
