# =============================================================================
# Hotel BI — Watcher Installer
# Run this ONCE to register the file watcher as a Windows Scheduled Task.
# After that, the watcher starts automatically every time you log in.
# =============================================================================

$taskName   = "HotelBI-GitWatcher"
$scriptPath = "C:\Users\v-krb\Claude Code Projects\Hotel BI template\watch-and-push.ps1"
$logPath    = "C:\Users\v-krb\Claude Code Projects\Hotel BI template\watcher.log"

# Remove existing task if it's already registered
if (Get-ScheduledTask -TaskName $taskName -ErrorAction SilentlyContinue) {
    Unregister-ScheduledTask -TaskName $taskName -Confirm:$false
    Write-Host "Removed existing task."
}

$action = New-ScheduledTaskAction `
    -Execute  "powershell.exe" `
    -Argument "-ExecutionPolicy Bypass -WindowStyle Hidden -NonInteractive -File `"$scriptPath`""

# Start automatically when you log in
$trigger = New-ScheduledTaskTrigger -AtLogOn

$settings = New-ScheduledTaskSettingsSet `
    -ExecutionTimeLimit  ([TimeSpan]::Zero) `
    -RestartCount        5 `
    -RestartInterval     (New-TimeSpan -Minutes 1) `
    -StartWhenAvailable

Register-ScheduledTask `
    -TaskName  $taskName `
    -Action    $action `
    -Trigger   $trigger `
    -Settings  $settings `
    -Force | Out-Null

Write-Host ""
Write-Host "Task '$taskName' registered successfully." -ForegroundColor Green
Write-Host "It will start automatically every time you log in to Windows."
Write-Host ""

# Start the task now so you don't need to log out and back in
Start-ScheduledTask -TaskName $taskName
Write-Host "Watcher is now running in the background." -ForegroundColor Green
Write-Host ""
Write-Host "Log file: $logPath"
Write-Host "To check activity, open the log file or run:"
Write-Host "  Get-Content '$logPath' -Tail 20"
Write-Host ""
Write-Host "To stop the watcher:  Stop-ScheduledTask  -TaskName '$taskName'"
Write-Host "To start it again:    Start-ScheduledTask -TaskName '$taskName'"
Write-Host "To uninstall it:      Unregister-ScheduledTask -TaskName '$taskName' -Confirm:`$false"
