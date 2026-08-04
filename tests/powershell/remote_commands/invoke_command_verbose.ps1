# vybe-test: powershell/remote_commands/invoke_command_verbose
$result = Invoke-Command -ScriptBlock { Write-Verbose 'v' } -Verbose
if ($result -ne $null) {
    Write-Host "FAIL: expected no direct output"
    exit 1
}
Write-Host "PASS"
exit 0
