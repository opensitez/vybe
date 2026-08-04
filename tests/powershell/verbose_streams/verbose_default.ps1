# vybe-test: powershell/verbose_streams/verbose_default
if ($VerbosePreference -ne 'SilentlyContinue') {
    Write-Host "FAIL: expected default SilentlyContinue"
    exit 1
}
Write-Host 'PASS'
exit 0
