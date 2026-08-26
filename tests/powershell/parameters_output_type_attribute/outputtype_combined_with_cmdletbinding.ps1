# vybe-test: powershell/parameters_output_type_attribute/outputtype_combined_with_cmdletbinding
function Test-OutputCmdlet {
    [OutputType([guid])]
    [CmdletBinding()]
    param()
    return [guid]::Empty
}
$cmd = Get-Command Test-OutputCmdlet
if ($cmd.OutputType[0].Type -ne [guid]) {
    Write-Host "FAIL: OutputType combined with CmdletBinding failed"
    exit 1
}
Write-Host "PASS"
exit 0
