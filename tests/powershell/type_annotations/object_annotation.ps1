# vybe-test: powershell/type_annotations/object_annotation
[object]$x = 'PASS'
if ($x -eq 'PASS') { Write-Host 'PASS'; exit 0 }
Write-Host 'FAIL'
exit 1
