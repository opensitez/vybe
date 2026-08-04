# vybe-test: powershell/parameter_sets/parameter_set_name
function Test-Func {
  [CmdletBinding(DefaultParameterSetName='A')]
  param([Parameter(ParameterSetName='A')][string]$A, [Parameter(ParameterSetName='B')][string]$B)
  return $PSCmdlet.ParameterSetName
}
if ((Test-Func -A 'x') -eq 'A') { Write-Host 'PASS'; exit 0 }
Write-Host 'FAIL'
exit 1
