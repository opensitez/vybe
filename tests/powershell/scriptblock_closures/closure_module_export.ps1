# vybe-test: powershell/scriptblock_closures/closure_module_export
$prefix = "MOD_PREFIX"
$exportedClosure = { param($t) "$prefix:$t" }.GetClosure()
$res = &$exportedClosure "ITEM"
if ($res -ne "MOD_PREFIX:ITEM") {
    Write-Host "FAIL: exported closure expected MOD_PREFIX:ITEM, got $res"
    exit 1
}
Write-Host "PASS"
exit 0
