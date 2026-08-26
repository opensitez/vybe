# vybe-test: powershell/classes_structural_equality_and_hash/structural_equality_check_16
class StructItem_16 {
    [int]$A = 16
    [string]$B = "Tag_16"
    [bool]Equals([object]$other) {
        if ($other -isnot [StructItem_16]) { return $false }
        return ($this.A -eq $other.A -and $this.B -eq $other.B)
    }
    [int]GetHashCode() {
        return [System.HashCode]::Combine($this.A, $this.B)
    }
}
$s1 = [StructItem_16]::new()
$s2 = [StructItem_16]::new()
if (-not $s1.Equals($s2) -or $s1.GetHashCode() -ne $s2.GetHashCode()) { Write-Host "FAIL: Structural equality failed"; exit 1 }
Write-Host "PASS"; exit 0
