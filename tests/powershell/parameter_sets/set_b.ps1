# vybe-test: powershell/parameter_sets/set_b
function Test-Func {
  [CmdletBinding(DefaultParameterSetName='A')]
  param([Parameter(ParameterSetName='A')][string]$A, [Parameter(ParameterSetName='B')][string]$B)
  return $PSCmdlet.ParameterSetName
}
if ((Test-Func -B 'x') -eq 'B') { Write-Host 'PASS'; exit 0 }
Write-Host 'FAIL'
exit 1
