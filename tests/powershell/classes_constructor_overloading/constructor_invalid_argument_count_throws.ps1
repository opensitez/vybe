# vybe-test: powershell/classes_constructor_overloading/constructor_invalid_argument_count_throws
class StrictTarget {
    [string]$Name
    StrictTarget([string]$n) { $this.Name = $n }
}
$caught = $false
try {
    $x = [StrictTarget]::new("A", "B")
} catch {
    $caught = $true
}
if (-not $caught) {
    Write-Host "FAIL: Expected constructor resolution error on wrong argument count"
    exit 1
}
Write-Host "PASS"
exit 0
