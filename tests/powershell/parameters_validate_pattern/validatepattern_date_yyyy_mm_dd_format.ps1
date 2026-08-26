# vybe-test: powershell/parameters_validate_pattern/validatepattern_date_yyyy_mm_dd_format
function Set-DateStr {
    param([ValidatePattern('^\d{4}-\d{2}-\d{2}$')][string]$Date)
    return $Date
}
$res = Set-DateStr -Date "2026-08-26"
if ($res -ne "2026-08-26") {
    Write-Host "FAIL: ValidatePattern date format failed"
    exit 1
}
Write-Host "PASS"
exit 0
