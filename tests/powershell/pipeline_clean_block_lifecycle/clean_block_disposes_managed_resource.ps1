# vybe-test: powershell/pipeline_clean_block_lifecycle/clean_block_disposes_managed_resource
class DisposableTracker : System.IDisposable {
    [bool]$Disposed = $false
    [void]Dispose() { $this.Disposed = $true }
}
$tracker = [DisposableTracker]::new()
function Test-ResourceClean {
    [CmdletBinding()]
    param([Parameter(ValueFromPipeline=$true)][int]$Val, [DisposableTracker]$Tracker)
    begin { $t = $Tracker }
    process { if ($Val -eq 5) { throw "Boom" } }
    clean { $t.Dispose() }
}
try {
    1, 5, 10 | Test-ResourceClean -Tracker $tracker
} catch {}
if (-not $tracker.Disposed) {
    Write-Host "FAIL: Clean block resource disposal failed"
    exit 1
}
Write-Host "PASS"
exit 0
