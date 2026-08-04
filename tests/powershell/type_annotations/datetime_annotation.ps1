# vybe-test: powershell/type_annotations/datetime_annotation
[datetime]$x = '2026-08-04'
if ($x.Year -eq 2026) { Write-Host 'PASS'; exit 0 }
Write-Host 'FAIL'
exit 1
