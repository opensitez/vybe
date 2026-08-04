# vybe-test: powershell/parameter_sets/pipeline_property_name
function Test-Func {
  [CmdletBinding(DefaultParameterSetName='A')]
  param([Parameter(ParameterSetName='A', ValueFromPipelineByPropertyName=$true)]$A, [Parameter(ParameterSetName='B', ValueFromPipelineByPropertyName=$true)]$B)
  return $PSCmdlet.ParameterSetName
}
$obj = [pscustomobject]@{ A = 'x' }
if (( $obj | Test-Func ) -eq 'A') { Write-Host 'PASS'; exit 0 }
Write-Host 'FAIL'
exit 1
