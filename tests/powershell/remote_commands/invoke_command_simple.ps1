# vybe-test: powershell/remote_commands/invoke_command_simple
$result = Invoke-Command -ScriptBlock { 4 + 4 }
if ($result -ne 8) {
    Write-Host "FAIL: expected 8, got $result"
    exit 1
}
Write-Host "PASS"
exit 0
