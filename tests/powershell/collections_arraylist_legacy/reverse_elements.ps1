# vybe-test: powershell/collections_arraylist_legacy/reverse_elements
$al = [System.Collections.ArrayList]::new()
$al.AddRange(@(1, 2, 3))
$al.Reverse()
if ($al[0] -ne 3 -or $al[1] -ne 2 -or $al[2] -ne 1) {
    Write-Host "FAIL: Reverse failed"
    exit 1
}
Write-Host "PASS"
exit 0
