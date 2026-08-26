# vybe-test: powershell/classes_interface_implementation/interface_method_throwing_exception
class BadDisposable : System.IDisposable {
    [void]Dispose() {
        throw [System.InvalidOperationException]::new("Cannot dispose")
    }
}
$bd = [BadDisposable]::new()
$caught = $false
try {
    $bd.Dispose()
} catch [System.InvalidOperationException] {
    $caught = $true
}
if (-not $caught) {
    Write-Host "FAIL: Interface method exception throw failed"
    exit 1
}
Write-Host "PASS"
exit 0
