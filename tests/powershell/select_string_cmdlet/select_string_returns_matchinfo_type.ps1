# vybe-test: powershell/select_string_cmdlet/select_string_returns_matchinfo_type
# Select-String emits instances of Microsoft.PowerShell.Commands.MatchInfo
$match = "search target found here" | Select-String -Pattern "target"

if ($null -eq $match) {
    Write-Host "FAIL: Select-String returned `$null"
    exit 1
}

if (-not ($match -is [Microsoft.PowerShell.Commands.MatchInfo])) {
    Write-Host "FAIL: expected MatchInfo object, got: $($match.GetType().FullName)"
    exit 1
}

Write-Host "PASS"
exit 0
