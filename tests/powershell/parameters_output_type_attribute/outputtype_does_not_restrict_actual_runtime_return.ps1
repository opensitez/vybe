# vybe-test: powershell/parameters_output_type_attribute/outputtype_does_not_restrict_actual_runtime_return
function Get-UntypedRuntime {
    [OutputType([int])]
    param()
    # OutputType is metadata only, PowerShell does not strictly cast returned objects
    return "returned-string"
}
$res = Get-UntypedRuntime
if ($res -ne "returned-string") {
    Write-Host "FAIL: OutputType should not alter runtime return value, got '$res'"
    exit 1
}
Write-Host "PASS"
exit 0
