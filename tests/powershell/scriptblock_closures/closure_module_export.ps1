# vybe-test: powershell/scriptblock_closures/closure_module_export
function Create-ExportClosure {
    $val = 99
    return { return $val * 2 }.GetNewClosure()
}
$c = Create-ExportClosure
$res = &$c
if ($res -ne 198) {
    Write-Host "FAIL: Closure module export failed"
    exit 1
}
Write-Host "PASS"
exit 0
