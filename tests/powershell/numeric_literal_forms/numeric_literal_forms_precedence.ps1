# vybe-test: powershell/numeric_literal_forms/precedence
if (2 + 3 * 4 - 5 -ne 9) {
  Write-Host 'FAIL: precedence evaluation mismatch'
  exit 1
}
Write-Host 'PASS'
exit 0
