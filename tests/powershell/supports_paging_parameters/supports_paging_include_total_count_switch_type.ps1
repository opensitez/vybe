# vybe-test: powershell/supports_paging_parameters/supports_paging_include_total_count_switch_type
# The IncludeTotalCount parameter added by SupportsPaging is a SwitchParameter
function TestIncludeTotalCountType {
    [CmdletBinding(SupportsPaging = $true)]
    param()
}

$paramInfo = (Get-Command TestIncludeTotalCountType).Parameters["IncludeTotalCount"]

if ($paramInfo.ParameterType -ne [System.Management.Automation.SwitchParameter]) {
    Write-Host "FAIL: expected IncludeTotalCount parameter type SwitchParameter, got $($paramInfo.ParameterType)"
    exit 1
}

if (-not $paramInfo.SwitchParameter) {
    Write-Host "FAIL: SwitchParameter property was false"
    exit 1
}

Write-Host "PASS"
exit 0
