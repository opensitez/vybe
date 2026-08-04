# vybe-test: powershell/parameter_sets/alone_set
function Test-Func {
  [CmdletBinding(DefaultParameterSetName='A')]
  param([Parameter(ParameterSetName='A')][string]$A)
  return $PSCmdlet.ParameterSetName
}
if ((Test-Func -A 'x') -eq 'A') { Write-Host 'PASS'; exit 0 }
Write-Host 'FAIL'
exit 1
