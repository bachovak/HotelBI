$scriptPath = "C:\Users\v-krb\Claude Code Projects\Hotel BI template\watch-and-push.ps1"

# Kill any existing watcher processes (PowerShell processes running the watcher script)
Get-WmiObject Win32_Process -Filter "Name='powershell.exe'" | ForEach-Object {
    if ($_.CommandLine -like "*watch-and-push*") {
        Stop-Process -Id $_.ProcessId -Force -ErrorAction SilentlyContinue
        Write-Host "Stopped existing watcher (PID $($_.ProcessId))."
    }
}

Start-Sleep -Seconds 2

# Start fresh
Start-Process powershell.exe `
    -ArgumentList "-ExecutionPolicy Bypass -WindowStyle Hidden -NonInteractive -File `"$scriptPath`"" `
    -WindowStyle Hidden

Write-Host "Watcher started."
