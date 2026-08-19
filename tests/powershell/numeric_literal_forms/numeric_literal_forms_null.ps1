# vybe-test: powershell/numeric_literal_forms/null
$val = $null
if ($val -ne $null) {
  Write-Host 'FAIL: null baseline should remain null'
  exit 1
}
Write-Host 'PASS'
exit 0
