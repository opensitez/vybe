# vybe-test: powershell/numeric_conversion_rules/invalid
try {
  [int]'not-a-number' | Out-Null
  Write-Host 'FAIL: invalid numeric conversion should throw'
  exit 1
} catch {
  Write-Host 'PASS'
  exit 0
}
