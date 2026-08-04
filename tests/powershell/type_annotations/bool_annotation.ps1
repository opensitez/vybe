# vybe-test: powershell/type_annotations/bool_annotation
[bool]$x = $true
if ($x) { Write-Host 'PASS'; exit 0 }
Write-Host 'FAIL'
exit 1
