# vybe-test: powershell/numeric_conversion_rules/edge
$val = [byte]255
if ($val -ne 255) {
  Write-Host "FAIL: byte conversion expected 255, got $val"
  exit 1
}
Write-Host 'PASS'
exit 0
