# vybe-test: powershell/parameter_sets/default_set
function Test-Func {
  [CmdletBinding(DefaultParameterSetName='A')]
  param([Parameter(ParameterSetName='A')][string]$A, [Parameter(ParameterSetName='B')][string]$B)
  return $PSCmdlet.ParameterSetName
}
if ((Test-Func) -eq 'A') { Write-Host 'PASS'; exit 0 }
Write-Host 'FAIL'
exit 1
