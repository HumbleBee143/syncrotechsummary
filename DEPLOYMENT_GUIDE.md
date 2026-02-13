# Deployment Guide

## 1. Choose the host machine
Use an always-on Windows PC/server with internet access and access to your publishing destination (SharePoint or internal web server).

## 2. Install prerequisites
- Git
- Windows PowerShell 5.1 (usually preinstalled)

## 3. Clone the repository
```powershell
cd C:\
mkdir Projects -Force
cd Projects
git clone https://github.com/HumbleBee143/syncrotechsummary
cd syncrotechsummary
```

## 4. Allow local scripts (if needed)
```powershell
Set-ExecutionPolicy -Scope CurrentUser RemoteSigned
```

## 5. Configure API key
Set the Syncro API key for the user account that will run the scheduled task:
```powershell
[Environment]::SetEnvironmentVariable("SYNCRO_API_KEY","YOUR_SYNCRO_API_KEY","User")
```
Close PowerShell and open a new window.

## 6. Verify API key is loaded
```powershell
echo $env:SYNCRO_API_KEY
```

## 7. Run once manually
```powershell
cd C:\Projects\syncrotechsummary
.\Run-SyncroTechSummary.ps1
```

## 8. Validate generated output
Confirm files exist in `output\`:
- `LatestReport.html`
- `Summary_*.html`
- `Open_*.html`
- `Closed_*.html`

## 9. Create scheduled task (every 30 minutes)
Run as current user:
```powershell
.\scripts\Register-SyncroTechSummaryTask.ps1 -TaskName "SyncroTechSummary" -IntervalMinutes 30 -UseCurrentUser
```

Run as service account:
```powershell
.\scripts\Register-SyncroTechSummaryTask.ps1 -TaskName "SyncroTechSummary" -IntervalMinutes 30 -RunAsUser "DOMAIN\svc-syncro"
```
You will be prompted for the password.

## 10. Test scheduled task immediately
```powershell
schtasks /Run /TN "SyncroTechSummary"
```
Check latest log in `output\Run_*.log`.

## 11. Publish for business access
Choose one:

- SharePoint document library:
  - Upload all files from `output\` to a dedicated folder.
  - Share link to `LatestReport.html`.

- Internal web server / IIS:
  - Point site/virtual directory at `C:\Projects\syncrotechsummary\output`.
  - Share URL to `LatestReport.html`.

## 12. Update process
When new changes are pushed:
```powershell
cd C:\Projects\syncrotechsummary
git pull
.\Run-SyncroTechSummary.ps1
```

## 13. Troubleshooting
- Script fails immediately:
  - Check `output\Run_*.log`.
- Unauthorized/API failures:
  - Re-set `SYNCRO_API_KEY` for the task-running user.
- Scheduled task runs but report not updated:
  - Confirm task account matches the account where `SYNCRO_API_KEY` is configured.
  - Confirm task action points to `Run-SyncroTechSummary.ps1` in this repo.
