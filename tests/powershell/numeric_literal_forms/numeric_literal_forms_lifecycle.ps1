# vybe-test: powershell/numeric_literal_forms/lifecycle
$values = @(1, 2, 3)
$values += 4
if ($values.Count -ne 4) {
  Write-Host "FAIL: lifecycle append changed count to $($values.Count)"
  exit 1
}
Write-Host 'PASS'
exit 0
