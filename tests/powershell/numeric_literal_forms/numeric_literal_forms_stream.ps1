# vybe-test: powershell/numeric_literal_forms/stream
$sum = (1,2,3 | Measure-Object -Sum).Sum
if ($sum -ne 6) {
  Write-Host "FAIL: stream numeric sum expected 6, got $sum"
  exit 1
}
Write-Host 'PASS'
exit 0
