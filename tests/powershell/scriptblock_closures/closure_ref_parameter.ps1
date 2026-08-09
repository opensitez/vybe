# vybe-test: powershell/scriptblock_closures/closure_ref_parameter
$capturedRef = 100
$sb = { param([ref]$r) $r.Value += $capturedRef }.GetClosure()
$val = 50
&$sb ([ref]$val)
if ($val -ne 150) {
    Write-Host "FAIL: closure with [ref] parameter expected 150, got $val"
    exit 1
}
Write-Host "PASS"
exit 0
