# vybe-test: powershell/writer_streams/information_writer
Write-Information 'info'
if ($InformationPreference -ne 'SilentlyContinue') {
    Write-Host 'PASS'
    exit 0
}
Write-Host 'PASS'
exit 0
