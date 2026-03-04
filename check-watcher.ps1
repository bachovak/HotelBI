$procs = Get-WmiObject Win32_Process -Filter "Name='powershell.exe'" |
         Where-Object { $_.CommandLine -like "*watch-and-push*" }

if ($procs) {
    foreach ($p in $procs) {
        $started = $p.ConvertToDateTime($p.CreationDate)
        Write-Host "Watcher is RUNNING (PID $($p.ProcessId), started $started)" -ForegroundColor Green
    }
} else {
    Write-Host "Watcher is NOT running." -ForegroundColor Red
    Write-Host "Start it with: Start-Process powershell.exe -ArgumentList '-ExecutionPolicy Bypass -WindowStyle Hidden -NonInteractive -File ""C:\Users\v-krb\Claude Code Projects\Hotel BI template\watch-and-push.ps1""' -WindowStyle Hidden"
}

Write-Host ""
Write-Host "Last 5 log entries:"
$log = "C:\Users\v-krb\Claude Code Projects\Hotel BI template\watcher.log"
if (Test-Path $log) {
    Get-Content $log -Tail 5
} else {
    Write-Host "(no log file yet)"
}
