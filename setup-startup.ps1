$scriptPath   = "C:\Users\v-krb\Claude Code Projects\Hotel BI template\watch-and-push.ps1"
$startupFolder = [Environment]::GetFolderPath('Startup')
$vbsPath      = Join-Path $startupFolder 'HotelBI-GitWatcher.vbs'

# VBScript launches PowerShell completely hidden (no window) at login
$vbs = @"
CreateObject("Wscript.Shell").Run "powershell.exe -ExecutionPolicy Bypass -WindowStyle Hidden -NonInteractive -File ""$scriptPath""", 0, False
"@

Set-Content -Path $vbsPath -Value $vbs -Encoding ASCII
Write-Host "Startup entry created: $vbsPath"

# Start the watcher right now without needing to log out
Start-Process powershell.exe `
    -ArgumentList "-ExecutionPolicy Bypass -WindowStyle Hidden -NonInteractive -File `"$scriptPath`"" `
    -WindowStyle Hidden

Write-Host "Watcher is now running in the background."
Write-Host "Log file: C:\Users\v-krb\Claude Code Projects\Hotel BI template\watcher.log"
