# vybe-test: powershell/information_streams/information_action_suppressed
$InformationPreference = 'SilentlyContinue'
Write-Information 'suppressed'
Write-Host 'PASS'
exit 0
