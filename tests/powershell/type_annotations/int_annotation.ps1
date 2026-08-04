# vybe-test: powershell/type_annotations/int_annotation
[int]$x = 5
if ($x -eq 5) { Write-Host 'PASS'; exit 0 }
Write-Host 'FAIL'
exit 1
