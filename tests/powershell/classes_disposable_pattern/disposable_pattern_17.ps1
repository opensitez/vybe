# vybe-test: powershell/classes_disposable_pattern/disposable_pattern_17
class ResourceHolder_17 : System.IDisposable {
    [bool]$IsDisposed = $false
    [void]Dispose() { $this.IsDisposed = $true }
}
$res = [ResourceHolder_17]::new()
$res.Dispose()
if (-not $res.IsDisposed) { Write-Host "FAIL: IDisposable pattern failed"; exit 1 }
Write-Host "PASS"; exit 0
