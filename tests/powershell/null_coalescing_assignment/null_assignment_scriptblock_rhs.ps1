# vybe-test: powershell/null_coalescing_assignment/null_assignment_scriptblock_rhs
$sb = $null
$sb ??= { param($x) $x * 2 }
$res = &$sb 25
if ($res -ne 50) {
    Write-Host "FAIL: scriptblock RHS ??= expected 50, got $res"
    exit 1
}
Write-Host "PASS"
exit 0
