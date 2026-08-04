# vybe-test: powershell/interpolation/dollar_brace
$msg = 'OK'
$text = "Status: ${msg}"
if ($text -ne 'Status: OK') {
    Write-Host 'FAIL'
    exit 1
}
Write-Host 'PASS'
exit 0
