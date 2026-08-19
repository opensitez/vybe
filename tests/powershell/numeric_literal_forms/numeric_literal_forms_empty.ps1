# vybe-test: powershell/numeric_literal_forms/empty
try {
  [int]'' | Out-Null
  Write-Host 'FAIL: empty numeric literal should throw'
  exit 1
} catch {
  Write-Host 'PASS'
  exit 0
}
