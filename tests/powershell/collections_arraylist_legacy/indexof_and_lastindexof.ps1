# vybe-test: powershell/collections_arraylist_legacy/indexof_and_lastindexof
$al = [System.Collections.ArrayList]::new()
$al.AddRange(@("x", "y", "x", "z"))
if ($al.IndexOf("x") -ne 0 -or $al.LastIndexOf("x") -ne 2) {
    Write-Host "FAIL: IndexOf / LastIndexOf failed"
    exit 1
}
Write-Host "PASS"
exit 0
