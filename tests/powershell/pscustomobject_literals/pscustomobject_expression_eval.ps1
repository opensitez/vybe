# vybe-test: powershell/pscustomobject_literals/pscustomobject_expression_eval
$a = 5
$b = 10
$obj = [pscustomobject]@{ Sum = $a + $b; Product = $a * $b }
if ($obj.Sum -ne 15 -or $obj.Product -ne 50) {
    Write-Host "FAIL: expression eval in literal expected Sum=15, Product=50"
    exit 1
}
Write-Host "PASS"
exit 0
