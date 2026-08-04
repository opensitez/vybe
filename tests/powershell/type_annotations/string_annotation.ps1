# vybe-test: powershell/type_annotations/string_annotation
[string]$x = 'PASS'
if ($x -eq 'PASS') { Write-Host 'PASS'; exit 0 }
Write-Host 'FAIL'
exit 1
