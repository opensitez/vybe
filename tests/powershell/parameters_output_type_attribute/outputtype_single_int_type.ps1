# vybe-test: powershell/parameters_output_type_attribute/outputtype_single_int_type
function Get-CounterVal {
    [OutputType([int])]
    param()
    return 42
}
$cmd = Get-Command Get-CounterVal
$outType = @($cmd.OutputType | ForEach-Object { $_.Type.Name })
if ($outType -notcontains "Int32") {
    Write-Host "FAIL: OutputType [int] metadata check failed"
    exit 1
}
Write-Host "PASS"
exit 0
