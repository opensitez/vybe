# vybe-test: powershell/collections_generic_hashset/exceptwith_set_difference
$s1 = [System.Collections.Generic.HashSet[string]]::new([string[]]@("a", "b", "c", "d"))
$s2 = [System.Collections.Generic.HashSet[string]]::new([string[]]@("b", "d"))
$s1.ExceptWith($s2)
if ($s1.Count -ne 2 -or -not $s1.Contains("a") -or -not $s1.Contains("c")) {
    Write-Host "FAIL: ExceptWith failed"
    exit 1
}
Write-Host "PASS"
exit 0
