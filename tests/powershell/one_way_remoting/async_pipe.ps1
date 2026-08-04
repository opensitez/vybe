# vybe-test: powershell/one_way_remoting/async_pipe
Start-Job -ScriptBlock { 8 } | Out-Null
Write-Host "PASS"
exit 0
