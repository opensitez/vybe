# vybe-test: powershell/collections_arraylist_legacy/add_heterogeneous_items
$al = [System.Collections.ArrayList]::new()
$al.Add(1)
$al.Add("text")
$al.Add($true)
if ($al.Count -ne 3 -or $al[0] -ne 1 -or $al[1] -ne "text" -or $al[2] -ne $true) {
    Write-Host "FAIL: Heterogeneous items Add failed"
    exit 1
}
Write-Host "PASS"
exit 0
