# vybe-test: powershell/parameters_output_type_attribute/outputtype_version_type
function Get-AppVer {
    [OutputType([version])]
    param()
    return [version]"1.0.0"
}
$cmd = Get-Command Get-AppVer
$types = @($cmd.OutputType | ForEach-Object { $_.Type.Name })
if ($types -notcontains "Version") {
    Write-Host "FAIL: OutputType Version check failed"
    exit 1
}
Write-Host "PASS"
exit 0
