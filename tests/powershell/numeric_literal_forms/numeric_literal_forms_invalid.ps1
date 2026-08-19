# vybe-test: powershell/numeric_literal_forms/invalid
try {
  [int]'abc' | Out-Null
  Write-Host 'FAIL: invalid numeric literal should throw'
  exit 1
} catch {
  Write-Host 'PASS'
  exit 0
}
