# vybe-test: powershell/collections_arraylist_legacy/clone_shallow_copy
$al1 = [System.Collections.ArrayList]::new()
$al1.Add("original")
$al2 = $al1.Clone()
$al1[0] = "mutated"
if ($al2[0] -ne "original") {
    Write-Host "FAIL: Clone should be shallow copy"
    exit 1
}
Write-Host "PASS"
exit 0
