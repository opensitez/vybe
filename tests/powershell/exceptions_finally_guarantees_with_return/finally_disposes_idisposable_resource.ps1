# vybe-test: powershell/exceptions_finally_guarantees_with_return/finally_disposes_idisposable_resource
class DisposableFileResource : System.IDisposable {
    [bool]$Disposed = $false
    [void]Dispose() { $this.Disposed = $true }
}
$res = [DisposableFileResource]::new()
function Use-Resource([DisposableFileResource]$r) {
    try {
        return "WorkComplete"
    } finally {
        $r.Dispose()
    }
}
$ret = Use-Resource $res
if ($ret -ne "WorkComplete" -or -not $res.Disposed) {
    Write-Host "FAIL: Resource disposal in finally failed"
    exit 1
}
Write-Host "PASS"
exit 0
