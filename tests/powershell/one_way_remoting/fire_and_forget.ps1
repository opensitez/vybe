# vybe-test: powershell/one_way_remoting/fire_and_forget
Start-Job -ScriptBlock { 6 } | Out-Null
Write-Host "PASS"
exit 0
