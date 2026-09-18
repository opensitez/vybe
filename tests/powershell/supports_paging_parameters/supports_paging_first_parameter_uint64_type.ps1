# vybe-test: powershell/supports_paging_parameters/supports_paging_first_parameter_uint64_type
# The First parameter added by SupportsPaging has type System.UInt64
function TestFirstParamType {
    [CmdletBinding(SupportsPaging = $true)]
    param()
}

$paramInfo = (Get-Command TestFirstParamType).Parameters["First"]

if ($paramInfo.ParameterType -ne [System.UInt64]) {
    Write-Host "FAIL: expected First parameter type UInt64, got $($paramInfo.ParameterType)"
    exit 1
}

Write-Host "PASS"
exit 0
