# vybe-test: powershell/arrays/negative_indexing
$arr = @(10, 20, 30, 40)
$last = $arr[-1]
if ($last -ne 40) {
    Write-Host "FAIL: expected 40, got $last"
    exit 1
}
$secondLast = $arr[-2]
if ($secondLast -ne 30) {
    Write-Host "FAIL: expected 30, got $secondLast"
    exit 1
}
Write-Host "PASS"
exit 0
