# vybe-test: powershell/invoke_expression_cmdlet/iex_dynamic_function_call
# Invoke-Expression can invoke functions defined in the current session by name
function DynCalcDouble { param($n) $n * 2 }

$result = Invoke-Expression "DynCalcDouble 7"

if ($result -ne 14) {
    Write-Host "FAIL: expected 14 from dynamic function call, got: $result"
    exit 1
}

Write-Host "PASS"
exit 0
