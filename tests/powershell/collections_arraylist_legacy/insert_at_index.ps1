# vybe-test: powershell/collections_arraylist_legacy/insert_at_index
$al = [System.Collections.ArrayList]::new()
$al.Add("A"); $al.Add("C")
$al.Insert(1, "B")
if ($al.Count -ne 3 -or $al[1] -ne "B") {
    Write-Host "FAIL: Insert failed"
    exit 1
}
Write-Host "PASS"
exit 0
