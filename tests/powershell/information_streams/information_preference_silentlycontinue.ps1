# vybe-test: powershell/information_streams/information_preference_silentlycontinue
$InformationPreference = 'SilentlyContinue'
Write-Information 'info'
Write-Host 'PASS'
exit 0
