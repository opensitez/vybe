# vybe-test: powershell/advanced_functions/parameter_sets
function Get-Value {
    [CmdletBinding(DefaultParameterSetName='ByName')]
    param(
        [Parameter(ParameterSetName='ByName')]
        [string]$Name,
        [Parameter(ParameterSetName='ById')]
        [int]$Id
    )
    if ($PSCmdlet.ParameterSetName -eq 'ByName') {
        return $Name
    }
    return $Id
}
$result = Get-Value -Name 'Alice'
if ($result -ne 'Alice') {
    Write-Host "FAIL: expected Alice, got $result"
    exit 1
}
Write-Host "PASS"
exit 0
