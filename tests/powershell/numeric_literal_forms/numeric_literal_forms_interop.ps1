# vybe-test: powershell/numeric_literal_forms/interop
if (([double]5).GetType().Name -ne 'Double') {
  Write-Host 'FAIL: interop typed literal should preserve double type'
  exit 1
}
Write-Host 'PASS'
exit 0
