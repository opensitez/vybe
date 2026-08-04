# vybe-test: powershell/functions/function_positional_parameters
function Multiply {
    param($x, $y)
    return $x * $y
}
$result = Multiply 6 7
if ($result -ne 42) {
    Write-Host "FAIL: expected 42, got $result"
    exit 1
}
Write-Host "PASS"
exit 0
