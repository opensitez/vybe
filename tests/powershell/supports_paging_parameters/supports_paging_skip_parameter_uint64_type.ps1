# vybe-test: powershell/supports_paging_parameters/supports_paging_skip_parameter_uint64_type
# The Skip parameter added by SupportsPaging has type System.UInt64
function TestSkipParamType {
    [CmdletBinding(SupportsPaging = $true)]
    param()
}

$paramInfo = (Get-Command TestSkipParamType).Parameters["Skip"]

if ($paramInfo.ParameterType -ne [System.UInt64]) {
    Write-Host "FAIL: expected Skip parameter type UInt64, got $($paramInfo.ParameterType)"
    exit 1
}

Write-Host "PASS"
exit 0
