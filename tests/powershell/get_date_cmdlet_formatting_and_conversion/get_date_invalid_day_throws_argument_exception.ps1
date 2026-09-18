# vybe-test: powershell/get_date_cmdlet_formatting_and_conversion/get_date_invalid_day_throws_argument_exception
# Passing an out-of-range day value (e.g. 32) to -Day throws an exception
$threwError = $false
try {
    Get-Date -Month 1 -Day 32 -ErrorAction Stop
} catch {
    $threwError = $true
}

if (-not $threwError) {
    Write-Host "FAIL: invalid day 32 did not throw an error"
    exit 1
}

Write-Host "PASS"
exit 0
