# vybe-test: powershell/object_sorting/sort_descending_numbers
$res = 5, 2, 8, 1, 9 | Sort-Object -Descending
if ($res[0] -ne 9 -or $res[4] -ne 1) {
    Write-Host "FAIL: Sort-Object -Descending numbers expected 9..1"
    exit 1
}
Write-Host "PASS"
exit 0
