# vybe-test: powershell/type_annotations/parameter_annotation
function Test-Func { param([int]$x); return $x }
if ((Test-Func -x 2) -eq 2) { Write-Host 'PASS'; exit 0 }
Write-Host 'FAIL'
exit 1
