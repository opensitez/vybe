# vybe-test: powershell/classes_interface_implementation/custom_equality_comparer_implementation
class CaseInsensitiveCustomComparer : System.Collections.Generic.IEqualityComparer[string] {
    [bool]Equals([string]$x, [string]$y) {
        return $x.Equals($y, [System.StringComparison]::OrdinalIgnoreCase)
    }
    [int]GetHashCode([string]$obj) {
        return $obj.ToLowerInvariant().GetHashCode()
    }
}
$comp = [CaseInsensitiveCustomComparer]::new()
$hs = [System.Collections.Generic.HashSet[string]]::new($comp)
$hs.Add("TEST")
if (-not $hs.Contains("test")) {
    Write-Host "FAIL: Custom IEqualityComparer in HashSet failed"
    exit 1
}
Write-Host "PASS"
exit 0
