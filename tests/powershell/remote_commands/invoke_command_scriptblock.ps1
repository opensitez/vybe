# vybe-test: powershell/remote_commands/invoke_command_scriptblock
$sb = { 'scriptblock' }
$result = Invoke-Command -ScriptBlock $sb
if ($result -ne 'scriptblock') {
    Write-Host "FAIL: expected scriptblock"
    exit 1
}
Write-Host "PASS"
exit 0
