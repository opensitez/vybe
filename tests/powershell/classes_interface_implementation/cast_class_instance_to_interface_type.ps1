# vybe-test: powershell/classes_interface_implementation/cast_class_instance_to_interface_type
class DisposableFile : System.IDisposable {
    [string]$Path
    [bool]$Closed = $false
    DisposableFile([string]$p) { $this.Path = $p }
    [void]Dispose() { $this.Closed = $true }
}
$df = [DisposableFile]::new("test.log")
$disp = [System.IDisposable]$df
$disp.Dispose()
if ($df.Closed -ne $true) {
    Write-Host "FAIL: Cast to interface and invoke method failed"
    exit 1
}
Write-Host "PASS"
exit 0
