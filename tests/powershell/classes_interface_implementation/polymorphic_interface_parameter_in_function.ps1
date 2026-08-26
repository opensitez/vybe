# vybe-test: powershell/classes_interface_implementation/polymorphic_interface_parameter_in_function
class Runner : System.IDisposable {
    [bool]$Executed = $false
    [void]Dispose() { $this.Executed = $true }
}
function Invoke-Cleanup([System.IDisposable]$d) {
    $d.Dispose()
}
$r = [Runner]::new()
Invoke-Cleanup $r
if (-not $r.Executed) {
    Write-Host "FAIL: Function accepting interface parameter failed"
    exit 1
}
Write-Host "PASS"
exit 0
