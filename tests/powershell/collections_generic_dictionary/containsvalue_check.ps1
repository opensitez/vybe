# vybe-test: powershell/collections_generic_dictionary/containsvalue_check
$d = [System.Collections.Generic.Dictionary[int, string]]::new()
$d.Add(1, "gold")
if (-not $d.ContainsValue("gold") -or $d.ContainsValue("silver")) {
    Write-Host "FAIL: ContainsValue failed"
    exit 1
}
Write-Host "PASS"
exit 0
