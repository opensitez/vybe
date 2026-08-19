# vybe-test: powershell/numeric_literal_forms/edge
$actual = 9223372036854775807
if ($actual -ne [long]::MaxValue) {
  Write-Host "FAIL: long max literal should match [long]::MaxValue"
  exit 1
}
Write-Host 'PASS'
exit 0
