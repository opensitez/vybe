# vybe-test: powershell/pipeline/sort_object_descending
$nums = @(3,1,4,1,5,9,2,6)
$sorted = $nums | Sort-Object -Descending
if ($sorted[0] -ne 9) { Write-Host "FAIL: first should be 9"; exit 1 }
if ($sorted[-1] -ne 1) { Write-Host "FAIL: last should be 1"; exit 1 }
Write-Host "PASS"
exit 0
