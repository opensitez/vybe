# vybe-test: powershell/classes_disposable_pattern/disposable_pattern_18
class ResourceHolder_18 : System.IDisposable {
    [bool]$IsDisposed = $false
    [void]Dispose() { $this.IsDisposed = $true }
}
$res = [ResourceHolder_18]::new()
$res.Dispose()
if (-not $res.IsDisposed) { Write-Host "FAIL: IDisposable pattern failed"; exit 1 }
Write-Host "PASS"; exit 0
