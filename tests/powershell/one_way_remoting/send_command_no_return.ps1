# vybe-test: powershell/one_way_remoting/send_command_no_return
$script = { Write-Output 'data' }
Invoke-Command -ScriptBlock $script | Out-Null
Write-Host "PASS"
exit 0
