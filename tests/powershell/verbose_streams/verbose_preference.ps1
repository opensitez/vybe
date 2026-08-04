# vybe-test: powershell/verbose_streams/verbose_preference
$VerbosePreference = 'Continue'
if ($VerbosePreference -ne 'Continue') {
    Write-Host "FAIL: expected Continue"
    exit 1
}
Write-Host 'PASS'
exit 0
