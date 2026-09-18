# vybe-test: powershell/get_date_cmdlet_formatting_and_conversion/get_date_invalid_month_throws_argument_exception
# Passing an out-of-range month value (e.g. 13) to -Month throws an exception
$threwError = $false
try {
    Get-Date -Month 13 -ErrorAction Stop
} catch {
    $threwError = $true
}

if (-not $threwError) {
    Write-Host "FAIL: invalid month 13 did not throw an error"
    exit 1
}

Write-Host "PASS"
exit 0
