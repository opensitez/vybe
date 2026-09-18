# vybe-test: powershell/invoke_expression_cmdlet/iex_dynamically_built_command_string
# Invoke-Expression enables building and executing a command string at runtime
$cmdletName = "Write-Output"
$arg = "dynamic_arg"
$result = Invoke-Expression "$cmdletName $arg"

if ($result -ne "dynamic_arg") {
    Write-Host "FAIL: expected 'dynamic_arg', got: '$result'"
    exit 1
}

Write-Host "PASS"
exit 0
