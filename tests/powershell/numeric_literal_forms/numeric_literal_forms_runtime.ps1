# vybe-test: powershell/numeric_literal_forms/runtime
$actual = (2 * 3) + 1
if ($actual -ne 7) {
  Write-Host "FAIL: runtime arithmetic mismatch, got $actual"
  exit 1
}
Write-Host 'PASS'
exit 0
