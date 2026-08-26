# vybe-test: powershell/parameters_output_type_attribute/outputtype_array_type
function Get-StringArray {
    [OutputType([string[]])]
    param()
    return @("a", "b", "c")
}
$cmd = Get-Command Get-StringArray
$type = $cmd.OutputType[0].Type
if (-not $type.IsArray) {
    Write-Host "FAIL: OutputType array check failed"
    exit 1
}
Write-Host "PASS"
exit 0
