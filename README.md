# Jain Sangh Directory

## Features
- Member lookup from `ajs.xlsx` loaded at app startup into memory for all users:
  - Membership Number is generated from Excel Column A + Column B (example: `A1`)
  - Name field shows searchable suggestions from Excel
  - Selecting a member auto-fills Name, Membership Number, Address, and Mobile Number (when available)
- Local SQLite database (`bookings.db`)
- Concurrency-safe writes using SQLite WAL + transactional insert + DB triggers
- Confirmation page with Save as PDF + Print buttons
- Unique booking number generated for every booking
- Sends confirmation email via TurboSMTP API to:
  - test9@gmail.com
  - test@shah.com
- Admin panel:
  - Full booking list with all columns
  - Filters for every column
  - Excel export of filtered data (includes booking number)

## Local Run
1. Create and activate virtual environment
2. Install packages:
   ```powershell
   pip install -r requirements.txt
   ```
3. Configure env vars (example):
   ```powershell
   $env:FLASK_SECRET_KEY="your-secret"
   $env:TURBO_SMTP_CONSUMER_KEY="your-consumer-key"
   $env:TURBO_SMTP_CONSUMER_SECRET="your-consumer-secret"
   $env:TURBO_SMTP_FROM="hello@your-company.com"
   ```
4. Run:
   ```powershell
   python app.py
   ```
5. Open: http://localhost:8000
6. Admin panel: http://localhost:8000/admin

## IIS Hosting (Windows Server)
1. Install IIS + CGI feature.
2. Install Python 3.11 (or your version) and add to PATH.
3. Install dependencies:
   ```powershell
   pip install -r requirements.txt
   pip install wfastcgi
   ```
4. Register `wfastcgi` once:
   ```powershell
   wfastcgi-enable
   ```
5. In `web.config`, update `scriptProcessor` with your real python and wfastcgi paths.
6. Create an IIS website/app pointing to this folder:
   - Physical path: `D:\PX\2026\P2026-03\VOTS\ajs`
7. In IIS Application Settings, set TurboSMTP environment variables (same as above).
8. Grant IIS app pool user write permission to this folder so SQLite can create/update `bookings.db`.
9. Restart IIS site.

## Files
- `app.py`: Flask app + DB logic + email sending
- `templates/admin.html`: admin list + filters + export
- `wsgi.py`: WSGI entrypoint for IIS
- `web.config`: IIS FastCGI configuration

- for launching for your jain sangh contact on chirayu@chirayusoftware.com
- `templates/`: HTML templates
- `static/style.css`: UI styles
