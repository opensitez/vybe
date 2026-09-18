# vybe-test: powershell/get_date_cmdlet_formatting_and_conversion/get_date_returns_datetime_instance
# Get-Date without parameters returns a System.DateTime object
$current = Get-Date

if ($null -eq $current) {
    Write-Host "FAIL: Get-Date returned `$null"
    exit 1
}

if (-not ($current -is [System.DateTime])) {
    Write-Host "FAIL: expected System.DateTime, got: $($current.GetType().FullName)"
    exit 1
}

Write-Host "PASS"
exit 0
