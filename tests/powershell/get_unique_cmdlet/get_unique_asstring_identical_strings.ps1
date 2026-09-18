# vybe-test: powershell/get_unique_cmdlet/get_unique_asstring_identical_strings
# Consecutive identical strings are collapsed to a single instance with -AsString
$items = @("production", "production", "staging")
$unique = @($items | Get-Unique -AsString)

if ($unique.Count -ne 2) {
    Write-Host "FAIL: expected 2 unique strings, got $($unique.Count)"
    exit 1
}

if ($unique[0] -ne "production" -or $unique[1] -ne "staging") {
    Write-Host "FAIL: unique string values mismatch: @($($unique -join ', '))"
    exit 1
}

Write-Host "PASS"
exit 0
