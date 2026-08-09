# vybe-test: powershell/scriptblock_closures/closure_hashtable_capture
$map = @{ Environment = "Production" }
$sb = { $map.Environment }.GetClosure()
$res = &$sb
if ($res -ne "Production") {
    Write-Host "FAIL: hashtable capture in closure expected 'Production', got '$res'"
    exit 1
}
Write-Host "PASS"
exit 0
