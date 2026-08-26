# vybe-test: powershell/classes_interface_implementation/implement_idisposable_dispose_method
class ManagedResource : System.IDisposable {
    [bool]$IsDisposed = $false
    [void]Dispose() {
        $this.IsDisposed = $true
    }
}
$res = [ManagedResource]::new()
$res.Dispose()
if ($res.IsDisposed -ne $true) {
    Write-Host "FAIL: IDisposable implementation failed"
    exit 1
}
Write-Host "PASS"
exit 0
