# SyncroTechSummary

## Structure
- `scripts/` PowerShell scripts
- `config/` runtime config
- `assets/` logos/images
- `output/` generated reports/logs

## Run
```powershell
.\Run-SyncroTechSummary.ps1
```

## Secrets
- Preferred: set `SYNCRO_API_KEY` in your environment.
- Optional local override file: `config\Syncro-TechSummary.config.local.json` (ignored by git).
- Template: `config\Syncro-TechSummary.config.local.example.json`.

Example (current PowerShell session only):
```powershell
$env:SYNCRO_API_KEY = "your-syncro-api-key"
```

## Task Scheduler
Command-based setup (recommended):
```powershell
.\scripts\Register-SyncroTechSummaryTask.ps1 -TaskName "SyncroTechSummary" -IntervalMinutes 30 -UseCurrentUser
```

Run as service account:
```powershell
.\scripts\Register-SyncroTechSummaryTask.ps1 -TaskName "SyncroTechSummary" -IntervalMinutes 30 -RunAsUser "DOMAIN\svc-syncro"
```
You will be prompted for the account password.

XML import/export setup:
1. Edit `deployment\SyncroTechSummary.TaskTemplate.xml`:
- set `<UserId>`
- set `<Arguments>` script path
- set `<WorkingDirectory>` repo path
2. Import:
```powershell
schtasks /Create /TN "SyncroTechSummary" /XML ".\deployment\SyncroTechSummary.TaskTemplate.xml" /F
```
3. Export from one machine to clone exact settings elsewhere:
```powershell
schtasks /Query /TN "SyncroTechSummary" /XML > .\deployment\SyncroTechSummary.TaskExport.xml
```
