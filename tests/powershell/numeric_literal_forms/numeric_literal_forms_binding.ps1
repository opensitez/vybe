# vybe-test: powershell/numeric_literal_forms/binding
$a = 7
if ($a -ne 7) {
  Write-Host "FAIL: bound number mismatch, got $a"
  exit 1
}
Write-Host 'PASS'
exit 0
