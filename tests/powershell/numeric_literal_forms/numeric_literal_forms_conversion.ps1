# vybe-test: powershell/numeric_literal_forms/conversion
$actual = [int]'42'
if ($actual -ne 42) {
  Write-Host "FAIL: string to int conversion failed, got $actual"
  exit 1
}
Write-Host 'PASS'
exit 0
