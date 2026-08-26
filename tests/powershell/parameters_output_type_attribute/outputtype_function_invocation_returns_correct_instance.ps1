# vybe-test: powershell/parameters_output_type_attribute/outputtype_function_invocation_returns_correct_instance
function New-GuidInstance {
    [OutputType([guid])]
    param()
    return [guid]::Parse("11111111-2222-3333-4444-555555555555")
}
$res = New-GuidInstance
if ($res -isnot [guid] -or $res.ToString() -ne "11111111-2222-3333-4444-555555555555") {
    Write-Host "FAIL: Function invocation return failed"
    exit 1
}
Write-Host "PASS"
exit 0
