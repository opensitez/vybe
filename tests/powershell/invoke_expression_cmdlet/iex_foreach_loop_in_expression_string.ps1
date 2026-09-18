# vybe-test: powershell/invoke_expression_cmdlet/iex_foreach_loop_in_expression_string
# Invoke-Expression can execute foreach loops embedded in a string, affecting caller scope
Invoke-Expression 'foreach ($n in 1..3) { $iexForeachSum += $n }'

if ($iexForeachSum -ne 6) {
    Write-Host "FAIL: expected sum 6 from foreach loop, got: $iexForeachSum"
    exit 1
}

Write-Host "PASS"
exit 0
