# vybe-test: powershell/one_way_remoting/background_action
Start-Job -ScriptBlock { 'background' } | Out-Null
Write-Host "PASS"
exit 0
