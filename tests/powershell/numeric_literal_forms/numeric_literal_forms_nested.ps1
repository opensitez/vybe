# vybe-test: powershell/numeric_literal_forms/nested
if ((1 + (2 + (3))) -ne 6) {
  Write-Host 'FAIL: nested expression numeric literal grouping mismatch'
  exit 1
}
Write-Host 'PASS'
exit 0
