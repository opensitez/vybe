# vybe-test: powershell/cmdlets/invoke_expression
$expr = "2 + 2 * 10"
$result = Invoke-Expression $expr
if ($result -ne 22) {
    Write-Host "FAIL: expected 22, got $result"
    exit 1
}
$funcExpr = 'function Square($n) { $n * $n }; Square 7'
$sq = Invoke-Expression $funcExpr
if ($sq -ne 49) {
    Write-Host "FAIL: expected 49, got $sq"
    exit 1
}
Write-Host "PASS"
exit 0
