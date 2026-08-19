# vybe-test: powershell/numeric_literal_forms/metadata
$meta = (Get-Date).ToString('yyyy')
if ($meta.Length -ne 4) {
  Write-Host "FAIL: date formatting metadata expected 4 digits, got '$meta'"
  exit 1
}
Write-Host 'PASS'
exit 0
