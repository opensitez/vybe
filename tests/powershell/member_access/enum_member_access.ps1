# vybe-test: powershell/member_access/enum_member_access
if ([System.DayOfWeek]::Friday -ne [System.DayOfWeek]::Friday) {
    Write-Host 'FAIL'
    exit 1
}
Write-Host 'PASS'
exit 0
