# vybe-test: powershell/verbose_streams/verbose_action_silentlycontinue
$VerbosePreference = 'SilentlyContinue'
Write-Verbose 'x'
if ($VerbosePreference -ne 'SilentlyContinue') {
    Write-Host "FAIL: expected SilentlyContinue"
    exit 1
}
Write-Host 'PASS'
exit 0
