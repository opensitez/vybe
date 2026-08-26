# vybe-test: powershell/parameters_output_type_attribute/outputtype_uri_type
function Get-ServiceUri {
    [OutputType([System.Uri])]
    param()
    return [System.Uri]::new("https://api.example.com")
}
$cmd = Get-Command Get-ServiceUri
$types = @($cmd.OutputType | ForEach-Object { $_.Type.Name })
if ($types -notcontains "Uri") {
    Write-Host "FAIL: OutputType Uri check failed"
    exit 1
}
Write-Host "PASS"
exit 0
