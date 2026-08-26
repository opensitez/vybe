# vybe-test: powershell/classes_interface_implementation/idisposable_in_using_object_simulation
class Cleanable : System.IDisposable {
    [bool]$Cleaned = $false
    [void]Dispose() { $this.Cleaned = $true }
}
$c = [Cleanable]::new()
try {
    # simulated using block
} finally {
    $c.Dispose()
}
if (-not $c.Cleaned) {
    Write-Host "FAIL: Simulated using block dispose failed"
    exit 1
}
Write-Host "PASS"
exit 0
