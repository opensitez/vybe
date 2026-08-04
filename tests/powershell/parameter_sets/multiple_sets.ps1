# vybe-test: powershell/parameter_sets/multiple_sets
function Test-Func {
  [CmdletBinding(DefaultParameterSetName='A')]
  param([Parameter(ParameterSetName='A')][string]$A, [Parameter(ParameterSetName='B')][string]$B, [Parameter(ParameterSetName='C')][string]$C)
  return $PSCmdlet.ParameterSetName
}
if ((Test-Func -C 'x') -eq 'C') { Write-Host 'PASS'; exit 0 }
Write-Host 'FAIL'
exit 1
