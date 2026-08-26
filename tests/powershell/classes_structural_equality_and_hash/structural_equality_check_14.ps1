# vybe-test: powershell/classes_structural_equality_and_hash/structural_equality_check_14
class StructItem_14 {
    [int]$A = 14
    [string]$B = "Tag_14"
    [bool]Equals([object]$other) {
        if ($other -isnot [StructItem_14]) { return $false }
        return ($this.A -eq $other.A -and $this.B -eq $other.B)
    }
    [int]GetHashCode() {
        return [System.HashCode]::Combine($this.A, $this.B)
    }
}
$s1 = [StructItem_14]::new()
$s2 = [StructItem_14]::new()
if (-not $s1.Equals($s2) -or $s1.GetHashCode() -ne $s2.GetHashCode()) { Write-Host "FAIL: Structural equality failed"; exit 1 }
Write-Host "PASS"; exit 0
