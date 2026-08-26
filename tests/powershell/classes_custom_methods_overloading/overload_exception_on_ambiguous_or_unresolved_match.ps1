# vybe-test: powershell/classes_custom_methods_overloading/overload_exception_on_ambiguous_or_unresolved_match
class StrictMethod {
    [void]Run([int]$x) {}
}
$sm = [StrictMethod]::new()
$caught = $false
try {
    $sm.Run("not", "matching", "args")
} catch {
    $caught = $true
}
if (-not $caught) {
    Write-Host "FAIL: Expected exception on invalid argument count"
    exit 1
}
Write-Host "PASS"
exit 0
