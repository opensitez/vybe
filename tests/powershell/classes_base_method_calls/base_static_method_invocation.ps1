# vybe-test: powershell/classes_base_method_calls/base_static_method_invocation
class BaseStatic {
    static [int]GetMultiplier() { return 5 }
}
class SubStatic : BaseStatic {
    static [int]Compute([int]$v) {
        return $v * [BaseStatic]::GetMultiplier()
    }
}
$res = [SubStatic]::Compute(10)
if ($res -ne 50) {
    Write-Host "FAIL: Base static method invocation failed, got $res"
    exit 1
}
Write-Host "PASS"
exit 0
