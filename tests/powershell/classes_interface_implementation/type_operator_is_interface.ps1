# vybe-test: powershell/classes_interface_implementation/type_operator_is_interface
class CustomDisposable : System.IDisposable {
    [void]Dispose() {}
}
$cd = [CustomDisposable]::new()
if ($cd -isnot [System.IDisposable]) {
    Write-Host "FAIL: -is [IDisposable] check failed"
    exit 1
}
Write-Host "PASS"
exit 0
