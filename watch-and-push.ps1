# Hotel BI - Auto Git Push Watcher
# Watches for .pbix changes and pushes to GitHub automatically.

$repoPath        = "C:\Users\v-krb\Claude Code Projects\Hotel BI template"
$watchPath       = "$repoPath\Hotel BI\Power BI"
$logFile         = "$repoPath\watcher.log"
$debounceSeconds = 30

function Write-Log {
    param([string]$msg)
    $ts = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    Add-Content -Path $logFile -Value "[$ts] $msg"
}

Write-Log "---"
Write-Log "Watcher started. Monitoring: $watchPath"

$watcher                     = New-Object System.IO.FileSystemWatcher
$watcher.Path                = $watchPath
$watcher.Filter              = "*.pbix"
$watcher.NotifyFilter        = [System.IO.NotifyFilters]::LastWrite
$watcher.EnableRaisingEvents = $true

while ($true) {
    $change = $watcher.WaitForChanged([System.IO.WatcherChangeTypes]::Changed, 10000)

    if ($change.TimedOut) { continue }

    $fileName = $change.Name
    Write-Log "Change detected: $fileName - waiting ${debounceSeconds}s for Power BI to finish saving..."
    Start-Sleep -Seconds $debounceSeconds

    # Drain any extra events that fired during the wait
    while (-not $watcher.WaitForChanged([System.IO.WatcherChangeTypes]::Changed, 500).TimedOut) {}

    try {
        Set-Location $repoPath
        git add "Hotel BI/Power BI/$fileName" 2>&1 | Out-Null

        $staged = git diff --cached --name-only 2>&1
        if (-not $staged) {
            Write-Log "No staged changes - skipping commit."
            continue
        }

        $ts        = Get-Date -Format "yyyy-MM-dd HH:mm"
        $commitMsg = "Update $fileName - $ts"
        git commit -m $commitMsg 2>&1 | Out-Null
        $pushOut = git push 2>&1

        if ($LASTEXITCODE -eq 0) {
            Write-Log "Pushed to GitHub successfully."
        } else {
            Write-Log "Push failed: $pushOut"
        }
    }
    catch {
        $errMsg = $_.Exception.Message
        Write-Log "Error during git operation: $errMsg"
    }
}
