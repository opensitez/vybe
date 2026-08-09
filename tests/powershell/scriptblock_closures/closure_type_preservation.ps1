# vybe-test: powershell/scriptblock_closures/closure_type_preservation
$dt = [datetime]"2026-01-01"
$sb = { $dt }.GetClosure()
$res = &$sb
if (-not ($res -is [datetime]) -or $res.Year -ne 2026) {
    Write-Host "FAIL: closure captured type preservation expected [datetime] Year=2026"
    exit 1
}
Write-Host "PASS"
exit 0
