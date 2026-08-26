# vybe-test: powershell/parameters_output_type_attribute/outputtype_boolean_type
function Test-ConnectionState {
    [OutputType([bool])]
    param()
    return $true
}
$cmd = Get-Command Test-ConnectionState
$types = @($cmd.OutputType | ForEach-Object { $_.Type.Name })
if ($types -notcontains "Boolean") {
    Write-Host "FAIL: OutputType Boolean check failed"
    exit 1
}
Write-Host "PASS"
exit 0
