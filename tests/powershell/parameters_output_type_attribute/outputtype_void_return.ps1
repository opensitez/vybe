# vybe-test: powershell/parameters_output_type_attribute/outputtype_void_return
function Invoke-VoidTask {
    [OutputType([void])]
    param()
}
$cmd = Get-Command Invoke-VoidTask
$types = @($cmd.OutputType | ForEach-Object { $_.Type.Name })
if ($types -notcontains "Void") {
    Write-Host "FAIL: OutputType Void check failed"
    exit 1
}
Write-Host "PASS"
exit 0
