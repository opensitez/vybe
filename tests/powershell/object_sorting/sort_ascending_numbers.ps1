# vybe-test: powershell/object_sorting/sort_ascending_numbers
$res = 5, 2, 8, 1, 9 | Sort-Object
if ($res[0] -ne 1 -or $res[4] -ne 9) {
    Write-Host "FAIL: Sort-Object ascending numbers expected 1..9"
    exit 1
}
Write-Host "PASS"
exit 0
