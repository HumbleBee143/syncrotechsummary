# Context

## Goal
Generate a recurring HTML technician report from Syncro API data, including weekly ticket activity, current open workload, and technician drilldowns suitable for internal business use.

## Current State
- Repository restructured for portability:
  - `scripts/` runnable PowerShell scripts
  - `config/` tracked config and local example
  - `assets/` logos/images
  - `output/` generated HTML/log/report artifacts
- Entry point:
  - `Run-SyncroTechSummary.ps1`
- Scheduling/deployment assets:
  - `scripts/Register-SyncroTechSummaryTask.ps1`
  - `deployment/SyncroTechSummary.TaskTemplate.xml`
  - `DEPLOYMENT_GUIDE.md`
- Security hardening completed:
  - API key is read from `SYNCRO_API_KEY` env var (preferred)
  - tracked config is sanitized
  - local secret config is git-ignored
  - history was rewritten to remove previously exposed secret content
- Reporting UI refreshed:
  - main summary page redesigned
  - technician open/closed pages redesigned
  - summary detail pages (`Summary_*.html`) redesigned
- Time tracking logic implemented/fixed:
  - `ticket_timers` parsing updated for Syncro fields (`active_duration`, `billable_time`, `start_time`, `end_time`, etc.)
  - timer paging updated to read latest pages (not only oldest)
  - non-zero weekly timer totals confirmed in latest runs
- Time display now includes:
  - per-ticket `Actual` time (within current report window)
  - per-ticket `Syncro Total` (from ticket `total_formatted_billable_time` when available)
  - per-tech weekly totals on technician pages
  - main page total card for actual time
  - main page per-tech actual-time graph section

## Known Constraints
- Some Syncro endpoints can return limited shapes depending on API key permissions.
- `/tickets/{id}` may not be accessible with all keys; report avoids hard dependency on that endpoint.
- If timer records lack user names, the script resolves names via `/api/v1/users` and falls back to `UserId:*` when needed.

## Run/Deploy Notes
- Config file:
  - `config/Syncro-TechSummary.config.json`
- Local secret override template:
  - `config/Syncro-TechSummary.config.local.example.json`
- Generated outputs:
  - `output/LatestReport.html`
  - `output/Summary_*.html`
  - `output/Open_*.html`
  - `output/Closed_*.html`
  - `output/Run_*.log`

## Suggested Next Steps
- Validate publishing workflow on target host (SharePoint or internal web path).
- Monitor scheduled runs for 1-2 weeks to confirm consistency.
- Optionally add alerting on run failures (email/Teams/webhook).
