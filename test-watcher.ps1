$path  = "C:\Users\v-krb\Claude Code Projects\Hotel BI template\Hotel BI\Power BI\Hotel_Report.pbix"
$bytes = [System.IO.File]::ReadAllBytes($path)
[System.IO.File]::WriteAllBytes($path, $bytes)
Write-Host "File written - watcher should fire in ~30 seconds."
Write-Host "Run this to check the log:"
Write-Host "  Get-Content 'C:\Users\v-krb\Claude Code Projects\Hotel BI template\watcher.log' -Tail 10"
