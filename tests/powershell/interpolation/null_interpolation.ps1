# vybe-test: powershell/interpolation/null_interpolation
$text = "Value: $($null)"
if ($text -ne 'Value: ') {
    Write-Host 'FAIL'
    exit 1
}
Write-Host 'PASS'
exit 0
