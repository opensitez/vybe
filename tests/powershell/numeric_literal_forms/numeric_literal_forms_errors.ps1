# vybe-test: powershell/numeric_literal_forms/errors
try {
  1 / 0 | Out-Null
  Write-Host 'FAIL: division by zero expected to error'
  exit 1
} catch {
  Write-Host 'PASS'
  exit 0
}
