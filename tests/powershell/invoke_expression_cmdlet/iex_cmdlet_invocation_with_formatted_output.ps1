# vybe-test: powershell/invoke_expression_cmdlet/iex_cmdlet_invocation_with_formatted_output
# Invoke-Expression can invoke cmdlets with parameters, including formatting flags like -Format
$result = Invoke-Expression "Get-Date -Year 2030 -Month 1 -Day 1 -Format 'yyyy'"

if ($result -ne "2030") {
    Write-Host "FAIL: expected '2030', got: '$result'"
    exit 1
}

Write-Host "PASS"
exit 0
