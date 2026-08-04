# vybe-test: powershell/information_streams/information_action_continue
$InformationPreference = 'Continue'
Write-Information 'x'
Write-Host 'PASS'
exit 0
