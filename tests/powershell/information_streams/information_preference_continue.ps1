# vybe-test: powershell/information_streams/information_preference_continue
$InformationPreference = 'Continue'
Write-Information 'info'
Write-Host 'PASS'
exit 0
