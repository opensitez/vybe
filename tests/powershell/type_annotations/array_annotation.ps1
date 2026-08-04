# vybe-test: powershell/type_annotations/array_annotation
[int[]]$x = 1,2,3
if ($x.Count -eq 3) { Write-Host 'PASS'; exit 0 }
Write-Host 'FAIL'
exit 1
