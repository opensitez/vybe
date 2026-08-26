# vybe-test: powershell/classes_polymorphism_and_casting/override_equals_and_gethashcode
class IdentItem {
    [int]$Id
    IdentItem([int]$i) { $this.Id = $i }
    [bool]Equals([object]$other) {
        if ($null -eq $other -or $other -isnot [IdentItem]) { return $false }
        return $this.Id -eq ([IdentItem]$other).Id
    }
    [int]GetHashCode() { return $this.Id.GetHashCode() }
}
$i1 = [IdentItem]::new(100)
$i2 = [IdentItem]::new(100)
$i3 = [IdentItem]::new(200)
if (-not $i1.Equals($i2) -or $i1.Equals($i3) -or $i1.GetHashCode() -ne $i2.GetHashCode()) {
    Write-Host "FAIL: Equals/GetHashCode override failed"
    exit 1
}
Write-Host "PASS"
exit 0
