# vybe-test: powershell/scriptblocks/scriptblock_closure_variable
$multiplier = 5
$sb = { param($n) $n * $multiplier }
$result = & $sb 8
if ($result -ne 40) {
    Write-Host "FAIL: expected 40, got $result"
    exit 1
}
Write-Host "PASS"
exit 0
