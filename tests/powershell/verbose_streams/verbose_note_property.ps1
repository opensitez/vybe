# vybe-test: powershell/verbose_streams/verbose_note_property
$VerbosePreference = 'Continue'
Write-Verbose 'v'
if ($VerbosePreference -ne 'Continue') {
    Write-Host "FAIL: expected continue preference"
    exit 1
}
Write-Host 'PASS'
exit 0
