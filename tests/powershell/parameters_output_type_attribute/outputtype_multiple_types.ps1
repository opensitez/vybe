# vybe-test: powershell/parameters_output_type_attribute/outputtype_multiple_types
function Get-MixedData {
    [OutputType([string], [int])]
    param([bool]$AsInt)
    if ($AsInt) { return 100 }
    return "one hundred"
}
$cmd = Get-Command Get-MixedData
$types = @($cmd.OutputType | ForEach-Object { $_.Type.Name })
if ($types -notcontains "String" -or $types -notcontains "Int32") {
    Write-Host "FAIL: OutputType multiple types metadata failed"
    exit 1
}
Write-Host "PASS"
exit 0
