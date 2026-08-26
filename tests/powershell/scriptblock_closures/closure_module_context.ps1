# vybe-test: powershell/scriptblock_closures/closure_module_context
function Create-ClosureContext {
    $secret = "ModSecret"
    return { return $secret }.GetNewClosure()
}
$c = Create-ClosureContext
$res = &$c
if ($res -ne "ModSecret") {
    Write-Host "FAIL: Closure module context failed"
    exit 1
}
Write-Host "PASS"
exit 0
