# vybe-test: powershell/automatic_variables/dollar_args_in_function
function SumAll {
    $total = 0
    foreach ($n in $args) { $total += $n }
    return $total
}
$result = SumAll 1 2 3 4 5
if ($result -ne 15) {
    Write-Host "FAIL: expected 15, got $result"
    exit 1
}
Write-Host "PASS"
exit 0
