# vybe-test: powershell/scriptblock_closures/closure_subexpression
$val = "ClosureInSubexpr"
$sb = { $val }.GetNewClosure()
$msg = "Result: $( &$sb )"
if ($msg -ne "Result: ClosureInSubexpr") {
    Write-Host "FAIL: closure execution in subexpression expected 'Result: ClosureInSubexpr', got '$msg'"
    exit 1
}
Write-Host "PASS"
exit 0
