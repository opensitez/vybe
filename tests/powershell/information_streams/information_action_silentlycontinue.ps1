# vybe-test: powershell/information_streams/information_action_silentlycontinue
$InformationPreference = 'SilentlyContinue'
Write-Information 'x'
Write-Host 'PASS'
exit 0
