# vybe-test: powershell/scriptblock_closures/closure_with_parameters
function Create-ParamClosure {
    $baseVal = 100
    return { param($extra) return $baseVal + $extra }.GetNewClosure()
}
$c = Create-ParamClosure
$res = & $c 50
if ($res -ne 150) {
    Write-Host "FAIL: Closure with parameters failed"
    exit 1
}
Write-Host "PASS"
exit 0
