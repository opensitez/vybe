# vybe-test: powershell/parameters_cmdletbinding_supports_should_process/defaultparametersetname_metadata_check
function Get-SetData {
    [CmdletBinding(DefaultParameterSetName="SetA")]
    param(
        [Parameter(ParameterSetName="SetA")][string]$A,
        [Parameter(ParameterSetName="SetB")][string]$B
    )
}
$cmd = Get-Command Get-SetData
if ($cmd.DefaultParameterSet -ne "SetA") {
    Write-Host "FAIL: DefaultParameterSetName metadata check failed"
    exit 1
}
Write-Host "PASS"
exit 0
