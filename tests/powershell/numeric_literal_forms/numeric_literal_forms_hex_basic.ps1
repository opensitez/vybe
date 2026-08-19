# vybe-test: powershell/numeric_literal_forms/hex_basic
$actual = 0xFF
if ($actual -ne 255) {
  Write-Host "FAIL: hex literal expected 255, got $actual"
  exit 1
}
Write-Host 'PASS'
exit 0
