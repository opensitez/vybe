# vybe-test: powershell/select_string_cmdlet/select_string_raw_parameter_returns_string
# In PowerShell v7+, -Raw emits raw string lines instead of MatchInfo wrappers
$lines = @("initial state", "target state reached", "final state")
$rawResult = $lines | Select-String -Pattern "target state" -Raw

if ($null -eq $rawResult) {
    Write-Host "FAIL: Select-String -Raw returned `$null"
    exit 1
}

if (-not ($rawResult -is [string])) {
    Write-Host "FAIL: expected [string] type with -Raw, got: $($rawResult.GetType().FullName)"
    exit 1
}

if ($rawResult -ne "target state reached") {
    Write-Host "FAIL: raw string mismatch, got: '$rawResult'"
    exit 1
}

Write-Host "PASS"
exit 0
