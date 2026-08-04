# vybe-test: powershell/one_way_remoting/remote_no_wait
Invoke-Command -ScriptBlock { 7 } | Out-Null
Write-Host "PASS"
exit 0
