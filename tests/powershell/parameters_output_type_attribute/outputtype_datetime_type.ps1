# vybe-test: powershell/parameters_output_type_attribute/outputtype_datetime_type
function Get-CurrentTimestamp {
    [OutputType([datetime])]
    param()
    return [datetime]::UtcNow
}
$cmd = Get-Command Get-CurrentTimestamp
$types = @($cmd.OutputType | ForEach-Object { $_.Type.Name })
if ($types -notcontains "DateTime") {
    Write-Host "FAIL: OutputType DateTime check failed"
    exit 1
}
Write-Host "PASS"
exit 0
