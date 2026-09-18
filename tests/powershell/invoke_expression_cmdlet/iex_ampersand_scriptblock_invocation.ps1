# vybe-test: powershell/invoke_expression_cmdlet/iex_ampersand_scriptblock_invocation
# Invoke-Expression can use & to invoke a scriptblock literal embedded in the string
$result = Invoke-Expression "& { 100 / 4 }"

if ($result -ne 25) {
    Write-Host "FAIL: expected 25 from scriptblock invocation, got: $result"
    exit 1
}

Write-Host "PASS"
exit 0
