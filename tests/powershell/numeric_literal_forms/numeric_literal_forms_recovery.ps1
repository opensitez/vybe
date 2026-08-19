# vybe-test: powershell/numeric_literal_forms/recovery
try {
  1 / 0 | Out-Null
} catch {
  $recovered = 5
}
if ($recovered -ne 5) {
  Write-Host 'FAIL: recovery path should set fallback numeric value'
  exit 1
}
Write-Host 'PASS'
exit 0
