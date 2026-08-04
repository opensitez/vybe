# vybe-test: powershell/remote_commands/invoke_command_error
$result = Invoke-Command -ScriptBlock { 1 + 1 }
if ($result -ne 2) {
    Write-Host "FAIL: expected 2"
    exit 1
}
Write-Host "PASS"
exit 0
