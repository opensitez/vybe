# vybe-test: powershell/classes_interface_implementation/interface_inheritance_in_derived_class
class BaseDisposable : System.IDisposable {
    [void]Dispose() {}
}
class DerivedDisposable : BaseDisposable {}
$dd = [DerivedDisposable]::new()
if ($dd -isnot [System.IDisposable]) {
    Write-Host "FAIL: Inherited interface conformance failed"
    exit 1
}
Write-Host "PASS"
exit 0
