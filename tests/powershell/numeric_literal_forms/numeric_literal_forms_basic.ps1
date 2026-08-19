# vybe-test: powershell/numeric_literal_forms/basic
$actual = 123
if ($actual -ne 123) {
  Write-Host "FAIL: basic integer literal should be 123, got $actual"
  exit 1
}
Write-Host 'PASS'
exit 0
