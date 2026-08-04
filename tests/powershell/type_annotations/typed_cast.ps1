# vybe-test: powershell/type_annotations/typed_cast
$x = [int]'5'
if ($x -eq 5) { Write-Host 'PASS'; exit 0 }
Write-Host 'FAIL'
exit 1
