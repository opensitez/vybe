# vybe-test: powershell/classes_interface_implementation/implement_iequatable_equals_method
class EquatableDemo {
    [int]$Id
    EquatableDemo([int]$i) { $this.Id = $i }
    [bool]Equals([object]$other) {
        if ($other -eq $null -or $other -isnot [EquatableDemo]) { return $false }
        return $this.Id -eq $other.Id
    }
}
$e1 = [EquatableDemo]::new(10)
$e2 = [EquatableDemo]::new(10)
$e3 = [EquatableDemo]::new(20)
if (-not $e1.Equals($e2) -or $e1.Equals($e3)) {
    Write-Host "FAIL: Equals method implementation failed"
    exit 1
}
Write-Host "PASS"
exit 0
