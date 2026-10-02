# Stop hook: warn when the repo has uncommitted or unpushed work.
$ErrorActionPreference = "SilentlyContinue"
Set-Location $env:CLAUDE_PROJECT_DIR
$changes = git status --porcelain
$ahead = git rev-list --count '@{u}..HEAD'
$msg = @()
if ($changes) { $msg += "$(@($changes).Count) uncommitted change(s)" }
if ([int]$ahead -gt 0) { $msg += "$ahead commit(s) not pushed" }
if ($msg.Count -gt 0) {
    @{ systemMessage = "GitHub reminder: " + ($msg -join ", ") + ". Ask Claude to commit + push." } | ConvertTo-Json -Compress
}
exit 0
