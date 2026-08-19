# vybe-test: powershell/numeric_literal_forms/scope
$outer = 1
$inner = & { 2 }
if (($outer + $inner) -ne 3) {
  Write-Host "FAIL: scoped numeric addition failed with outer=$outer inner=$inner"
  exit 1
}
Write-Host 'PASS'
exit 0
