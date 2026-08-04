# vybe-test: powershell/one_way_remoting/one_way_message
Invoke-Command -ScriptBlock { 'one-way' } | Out-Null
Write-Host "PASS"
exit 0
