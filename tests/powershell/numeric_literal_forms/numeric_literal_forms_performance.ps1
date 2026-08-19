# vybe-test: powershell/numeric_literal_forms/performance
$acc = 0
for ($i = 1; $i -le 1000; $i++) {
  $acc += $i
}
if ($acc -ne 500500) {
  Write-Host "FAIL: arithmetic loop expected 500500, got $acc"
  exit 1
}
Write-Host 'PASS'
exit 0
