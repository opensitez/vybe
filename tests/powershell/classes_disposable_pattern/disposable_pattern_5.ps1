# vybe-test: powershell/classes_disposable_pattern/disposable_pattern_5
class ResourceHolder_5 : System.IDisposable {
    [bool]$IsDisposed = $false
    [void]Dispose() { $this.IsDisposed = $true }
}
$res = [ResourceHolder_5]::new()
$res.Dispose()
if (-not $res.IsDisposed) { Write-Host "FAIL: IDisposable pattern failed"; exit 1 }
Write-Host "PASS"; exit 0
