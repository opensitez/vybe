# vybe-test: powershell/language_null_coalescing_and_assignment/null_coalescing_chained_multiple_operands
$a = $null
$b = $null
$c = "ThirdValue"
$d = "FourthValue"
$res = $a ?? $b ?? $c ?? $d
if ($res -ne "ThirdValue") {
    Write-Host "FAIL: Chained null coalescing failed, got '$res'"
    exit 1
}
Write-Host "PASS"
exit 0
