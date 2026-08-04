# vybe-test: powershell/type_annotations/decimal_annotation
[decimal]$x = 1.5
if ($x -eq 1.5) { Write-Host 'PASS'; exit 0 }
Write-Host 'FAIL'
exit 1
