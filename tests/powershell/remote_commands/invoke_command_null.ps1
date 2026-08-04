# vybe-test: powershell/remote_commands/invoke_command_null
$result = Invoke-Command -ScriptBlock { $null }
if ($result -ne $null) {
    Write-Host "FAIL: expected null"
    exit 1
}
Write-Host "PASS"
exit 0
