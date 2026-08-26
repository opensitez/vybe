# vybe-test: powershell/classes_polymorphism_and_casting/object_gettype_returns_actual_runtime_type
class BaseRuntime {}
class DerivedRuntime : BaseRuntime {}
[BaseRuntime]$inst = [DerivedRuntime]::new()
if ($inst.GetType().Name -ne "DerivedRuntime") {
    Write-Host "FAIL: GetType() should return runtime type 'DerivedRuntime', got '$($inst.GetType().Name)'"
    exit 1
}
Write-Host "PASS"
exit 0
