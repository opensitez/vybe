# vybe-test: powershell/classes_base_method_calls/base_method_override_polymorphism
class BasePoly {
    [string]GetRole() { return "Base" }
}
class DerivedPoly : BasePoly {
    [string]GetRole() { return "Derived" }
}
[BasePoly]$b = [DerivedPoly]::new()
if ($b.GetRole() -ne "Derived") {
    Write-Host "FAIL: Virtual dispatch polymorphism check failed, got '$($b.GetRole())'"
    exit 1
}
Write-Host "PASS"
exit 0
