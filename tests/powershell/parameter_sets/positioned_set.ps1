# vybe-test: powershell/parameter_sets/positioned_set
function Test-Func {
  [CmdletBinding(DefaultParameterSetName='A')]
  param([Parameter(ParameterSetName='A', Position=0)][string]$A, [Parameter(ParameterSetName='B', Position=0)][string]$B)
  return $PSCmdlet.ParameterSetName
}
if ((Test-Func 'x') -eq 'A') { Write-Host 'PASS'; exit 0 }
Write-Host 'FAIL'
exit 1
