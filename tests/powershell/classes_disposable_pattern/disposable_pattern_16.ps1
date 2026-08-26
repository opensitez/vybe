# vybe-test: powershell/classes_disposable_pattern/disposable_pattern_16
class ResourceHolder_16 : System.IDisposable {
    [bool]$IsDisposed = $false
    [void]Dispose() { $this.IsDisposed = $true }
}
$res = [ResourceHolder_16]::new()
$res.Dispose()
if (-not $res.IsDisposed) { Write-Host "FAIL: IDisposable pattern failed"; exit 1 }
Write-Host "PASS"; exit 0
