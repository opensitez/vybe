# vybe-test: powershell/verbose_streams/verbose_change
$VerbosePreference = 'Continue'
if ($VerbosePreference -ne 'Continue') {
    Write-Host "FAIL: expected changed preference"
    exit 1
}
Write-Host 'PASS'
exit 0
