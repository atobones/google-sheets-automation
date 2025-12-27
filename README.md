# Google Sheets Automation (Apps Script)

Practical Google Apps Script templates for automating Google Sheets.
Focused on small business workflows: leads, statuses, reporting.

## What’s inside
- Lead/CRM-style sheet with unique IDs
- Status workflow (NEW → IN_PROGRESS → DONE)
- Weekly report generation
- Archiving completed records
- Action logs for transparency

## Core features
- Auto-create base sheets and headers
- Add leads programmatically or manually
- Update lead status by ID
- Generate weekly status reports
- Archive completed leads
- Custom menu inside Google Sheets

## Typical use cases
- Simple CRM in Google Sheets
- Lead collection from forms or bots
- Internal request tracking
- Reporting automation
- Small business workflows

## Tech stack
- Google Apps Script
- Google Sheets

## How to run
1. Open a Google Sheet
2. Extensions → Apps Script
3. Paste code from `Code.gs`
4. Run `setupSheets()` once and allow permissions
5. Use the **Automation** menu inside the sheet

## Notes
This repository contains reusable automation templates.
The scripts are designed to be easily extended for client-specific needs
(e.g. webhooks, triggers, external integrations).
