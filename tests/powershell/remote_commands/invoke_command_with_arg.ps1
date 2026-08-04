# vybe-test: powershell/remote_commands/invoke_command_with_arg
$script = { param($n) $n * 3 }
$result = Invoke-Command -ScriptBlock $script -ArgumentList 4
if ($result -ne 12) {
    Write-Host "FAIL: expected 12, got $result"
    exit 1
}
Write-Host "PASS"
exit 0
