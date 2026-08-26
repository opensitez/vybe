# vybe-test: powershell/classes_base_method_calls/base_method_exception_propagation
class StrictBase {
    [void]Validate([int]$x) {
        if ($x -lt 0) { throw [System.ArgumentOutOfRangeException]::new("x must be positive") }
    }
}
class StrictSub : StrictBase {
    [void]Process([int]$x) {
        ([StrictBase]$this).Validate($x)
    }
}
$ss = [StrictSub]::new()
$caught = $false
try {
    $ss.Process(-10)
} catch [System.ArgumentOutOfRangeException] {
    $caught = $true
}
if (-not $caught) {
    Write-Host "FAIL: Base method exception propagation failed"
    exit 1
}
Write-Host "PASS"
exit 0
