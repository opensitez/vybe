# vybe-test: powershell/get_date_cmdlet_formatting_and_conversion/get_date_mutually_exclusive_format_and_uformat_throws
# Specifying both -Format and -UFormat violates parameter set exclusivity and throws a binding exception
$threwError = $false
try {
    Get-Date -Format "yyyy" -UFormat "%Y" -ErrorAction Stop
} catch {
    $threwError = $true
}

if (-not $threwError) {
    Write-Host "FAIL: combining -Format and -UFormat did not throw parameter binding error"
    exit 1
}

Write-Host "PASS"
exit 0
