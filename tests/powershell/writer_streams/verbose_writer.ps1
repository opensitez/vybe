# vybe-test: powershell/writer_streams/verbose_writer
$VerbosePreference = 'Continue'
Write-Verbose 'verb'
if ($VerbosePreference -ne 'Continue') {
    Write-Host "FAIL: expected verbose continue"
    exit 1
}
Write-Host 'PASS'
exit 0
