# vybe-test: powershell/remote_commands/invoke_command_complex
$result = Invoke-Command -ScriptBlock { @{ A = 1; B = 2 } }
if ($result['A'] -ne 1 -or $result['B'] -ne 2) {
    Write-Host "FAIL: expected hashtable values"
    exit 1
}
Write-Host "PASS"
exit 0
