# vybe-test: powershell/numeric_conversion_rules/basic
$val = [int]'15'
if ($val -ne 15) {
  Write-Host "FAIL: string to int conversion expected 15, got $val"
  exit 1
}
Write-Host 'PASS'
exit 0
