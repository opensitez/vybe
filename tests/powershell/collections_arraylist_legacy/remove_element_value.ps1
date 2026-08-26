# vybe-test: powershell/collections_arraylist_legacy/remove_element_value
$al = [System.Collections.ArrayList]::new()
$al.Add("keep"); $al.Add("del"); $al.Add("keep2")
$al.Remove("del")
if ($al.Count -ne 2 -or $al.Contains("del")) {
    Write-Host "FAIL: Remove failed"
    exit 1
}
Write-Host "PASS"
exit 0
