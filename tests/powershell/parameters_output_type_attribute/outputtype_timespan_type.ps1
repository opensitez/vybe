# vybe-test: powershell/parameters_output_type_attribute/outputtype_timespan_type
function Get-Elapsed {
    [OutputType([timespan])]
    param()
    return [timespan]::FromSeconds(5)
}
$cmd = Get-Command Get-Elapsed
$types = @($cmd.OutputType | ForEach-Object { $_.Type.Name })
if ($types -notcontains "TimeSpan") {
    Write-Host "FAIL: OutputType TimeSpan check failed"
    exit 1
}
Write-Host "PASS"
exit 0
