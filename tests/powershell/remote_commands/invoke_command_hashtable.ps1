# vybe-test: powershell/remote_commands/invoke_command_hashtable
$result = Invoke-Command -ScriptBlock { [hashtable]@{ X = 5 } }
if ($result['X'] -ne 5) {
    Write-Host "FAIL: expected hashtable value 5"
    exit 1
}
Write-Host "PASS"
exit 0
