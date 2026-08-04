# vybe-test: powershell/remote_commands/invoke_command_return_array
$result = Invoke-Command -ScriptBlock { 1,2,3 }
if ($result.Count -ne 3) {
    Write-Host "FAIL: expected 3 items"
    exit 1
}
Write-Host "PASS"
exit 0
