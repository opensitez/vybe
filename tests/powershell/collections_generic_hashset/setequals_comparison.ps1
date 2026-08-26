# vybe-test: powershell/collections_generic_hashset/setequals_comparison
$s1 = [System.Collections.Generic.HashSet[string]]::new([string[]]@("x", "y", "z"))
$s2 = [System.Collections.Generic.HashSet[string]]::new([string[]]@("z", "x", "y"))
if (-not $s1.SetEquals($s2)) {
    Write-Host "FAIL: SetEquals failed"
    exit 1
}
Write-Host "PASS"
exit 0
