# vybe-test: powershell/parameters_output_type_attribute/outputtype_single_string_type
function Get-StringOutput {
    [OutputType([string])]
    param()
    return "Hello"
}
$cmd = Get-Command Get-StringOutput
$outType = @($cmd.OutputType | ForEach-Object { $_.Type.Name })
if ($outType -notcontains "String") {
    Write-Host "FAIL: OutputType [string] metadata check failed"
    exit 1
}
Write-Host "PASS"
exit 0
