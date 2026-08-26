# vybe-test: powershell/parameters_output_type_attribute/outputtype_hashtable_type
function Get-ConfigHtOut {
    [OutputType([hashtable])]
    param()
    return @{ status = "OK" }
}
$cmd = Get-Command Get-ConfigHtOut
$types = @($cmd.OutputType | ForEach-Object { $_.Type.Name })
if ($types -notcontains "Hashtable") {
    Write-Host "FAIL: OutputType hashtable check failed"
    exit 1
}
Write-Host "PASS"
exit 0
