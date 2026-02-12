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
