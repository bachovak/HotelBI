# =============================================================================
# Hotel BI — Auto Git Push Watcher
# Watches for .pbix changes and pushes to GitHub automatically.
# This script is managed by a Windows Scheduled Task (see install-watcher.ps1).
# =============================================================================

$repoPath  = "C:\Users\v-krb\Claude Code Projects\Hotel BI template"
$watchPath = "$repoPath\Hotel BI\Power BI"
$logFile   = "$repoPath\watcher.log"

# How long to wait after a change is detected before committing.
# Power BI Desktop writes the file in stages, so we give it time to finish.
$debounceSeconds = 30

function Write-Log {
    param([string]$Message)
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    Add-Content -Path $logFile -Value "[$timestamp] $Message"
}

Write-Log "---"
Write-Log "Watcher started. Monitoring: $watchPath"

$watcher = New-Object System.IO.FileSystemWatcher
$watcher.Path                = $watchPath
$watcher.Filter              = "*.pbix"
$watcher.NotifyFilter        = [System.IO.NotifyFilters]::LastWrite
$watcher.EnableRaisingEvents = $true

while ($true) {
    # Block here until a .pbix file changes (check every 10 seconds)
    $change = $watcher.WaitForChanged([System.IO.WatcherChangeTypes]::Changed, 10000)

    if ($change.TimedOut) { continue }

    $fileName = $change.Name
    Write-Log "Change detected: $fileName — waiting ${debounceSeconds}s for Power BI to finish saving..."

    Start-Sleep -Seconds $debounceSeconds

    # Drain any extra events that fired during the wait
    while (-not $watcher.WaitForChanged([System.IO.WatcherChangeTypes]::Changed, 500).TimedOut) {}

    try {
        Set-Location $repoPath

        # Stage the changed file
        $gitAdd = git add "Hotel BI/Power BI/$fileName" 2>&1

        # Only commit if there is actually something staged
        $staged = git diff --cached --name-only 2>&1
        if (-not $staged) {
            Write-Log "No staged changes — skipping commit."
            continue
        }

        $timestamp  = Get-Date -Format "yyyy-MM-dd HH:mm"
        $gitCommit  = git commit -m "Update $fileName — $timestamp" 2>&1
        $gitPush    = git push 2>&1

        if ($LASTEXITCODE -eq 0) {
            Write-Log "Pushed to GitHub successfully."
        } else {
            Write-Log "Push failed. Git output: $gitPush"
        }
    }
    catch {
        Write-Log "Error during git operation: $_"
    }
}
