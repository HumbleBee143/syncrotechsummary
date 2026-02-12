# Context

## Goal
Weekly Syncro technician summary for Bigfoot Networks with last-work-week stats, current open ticket views, and per-tech drilldowns for team meetings.

## Last State
- Report window set to last full work week (Mon–Fri local time).
- Main HTML report (LatestReport.html) includes:
  - Last Week and Current summary card sections (all cards clickable to summary pages).
  - Per-tech open tickets bar chart with clickable tech chips and rows.
  - Per-tech closed tickets (last work week) bar chart with clickable tech chips and rows.
  - Tech colors applied to chips, bars, and grouped ticket lists.
- Per-tech open pages:
  - Status columns with colored headers and light-blue column backgrounds.
  - Tickets grouped by status with status-colored borders and Syncro ticket links.
  - Header card is light blue with tech-colored title strip.
  - Info line: "Open: X | 14+ day open: Y" in blue inline block.
  - "Return to Summary" button styled as a button.
- Per-tech closed pages:
  - Header card is light blue with tech-colored title strip.
  - Week + total closed info in blue inline blocks.
  - Closed tickets list with Syncro links; status color for Resolved.
  - "Return to Summary" button styled as a button (same as open pages).
- Logo: BIGFOOT_WHITE_B200.png used in header; background color set to #0077c0 across all pages.
- Status colors added for key statuses; open statuses expanded to include Waiting on Supplier/Parts, Scheduled, Escalation.

## Next Steps
- Decide deployment approach (SharePoint recommended for a live HTML page).
- Optionally add automatic email delivery (link or attachment).
- Validate UI on a few tech pages and adjust spacing/contrast if needed.

## Notes
- Config: config\Syncro-TechSummary.config.json (OpenTickets.Statuses expanded).
- Reports output: output\LatestReport.html (or Output.ReportPath in config) and per-tech pages (Open_*.html, Closed_*.html).
- Logo copied to output\logo.png during report generation.
