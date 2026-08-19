# vybe-test: powershell/unary_expression_rules/basic
$base = 7
$neg = -$base
$restored = -$neg

if ($neg -ne -7 -or $restored -ne 7) {
    Write-Host "FAIL: unary sign handling unexpected: neg=$neg restored=$restored"
    exit 1
}

$flag = -$false
if ($flag -ne 0) {
    Write-Host "FAIL: unary minus on boolean should yield 0, got $flag"
    exit 1
}

Write-Host 'PASS'
exit 0
