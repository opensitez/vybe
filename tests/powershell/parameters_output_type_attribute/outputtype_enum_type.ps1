# vybe-test: powershell/parameters_output_type_attribute/outputtype_enum_type
enum TestStatus { Pending; Running; Done }
function Get-StatusEnum {
    [OutputType([TestStatus])]
    param()
    return [TestStatus]::Done
}
$cmd = Get-Command Get-StatusEnum
$types = @($cmd.OutputType | ForEach-Object { $_.Type.Name })
if ($types -notcontains "TestStatus") {
    Write-Host "FAIL: OutputType enum check failed"
    exit 1
}
Write-Host "PASS"
exit 0
