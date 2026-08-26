# vybe-test: powershell/parameters_output_type_attribute/outputtype_generic_list_type
function Test-OutputList {
    [OutputType([System.Collections.Generic.List[string]])]
    [CmdletBinding()]
    param()
    return [System.Collections.Generic.List[string]]::new()
}
$cmd = Get-Command Test-OutputList
if ($cmd.OutputType[0].Type -ne [System.Collections.Generic.List[string]]) {
    Write-Host "FAIL: OutputType generic list check failed"
    exit 1
}
Write-Host "PASS"
exit 0
