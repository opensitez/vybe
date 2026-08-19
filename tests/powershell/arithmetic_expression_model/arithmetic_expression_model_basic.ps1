# vybe-test: powershell/arithmetic_expression_model/basic
$result = (10 + 2) * 3 - 4
if ($result -ne 32) {
    Write-Host "FAIL: expected 32, got $result"
    exit 1
}

$div = 20 / (2 + 3)
if ($div -ne 4) {
    Write-Host "FAIL: expected 4, got $div"
    exit 1
}

Write-Host 'PASS'
exit 0
