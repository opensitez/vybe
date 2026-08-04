# vybe-test: powershell/interpolation/escaped_dollar
$text = "Cost: `$5"
if ($text -ne 'Cost: $5') {
    Write-Host 'FAIL'
    exit 1
}
Write-Host 'PASS'
exit 0
