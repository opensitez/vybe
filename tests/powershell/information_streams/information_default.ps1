# vybe-test: powershell/information_streams/information_default
if ($InformationPreference -ne 'SilentlyContinue') {
    Write-Host 'FAIL: expected default SilentlyContinue'
    exit 1
}
Write-Host 'PASS'
exit 0
