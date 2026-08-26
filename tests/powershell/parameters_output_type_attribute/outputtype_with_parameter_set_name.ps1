# vybe-test: powershell/parameters_output_type_attribute/outputtype_with_parameter_set_name
function Test-OutputSet {
    [OutputType([string], ParameterSetName="StringMode")]
    [OutputType([int], ParameterSetName="IntMode")]
    [CmdletBinding()]
    param(
        [Parameter(ParameterSetName="StringMode")][switch]$AsString,
        [Parameter(ParameterSetName="IntMode")][switch]$AsInt
    )
    return "ok"
}
$cmd = Get-Command Test-OutputSet
if ($cmd.OutputType.Count -lt 2) {
    Write-Host "FAIL: OutputType with parameter set name failed"
    exit 1
}
Write-Host "PASS"
exit 0
