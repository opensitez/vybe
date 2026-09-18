# vybe-test: powershell/supports_paging_parameters/supports_paging_include_total_count_flag_present
# Passing -IncludeTotalCount sets $PSCmdlet.PagingParameters.IncludeTotalCount to true
function TestIncludeTotalBinding {
    [CmdletBinding(SupportsPaging = $true)]
    param()
    process {
        return $PSCmdlet.PagingParameters.IncludeTotalCount.IsPresent
    }
}

$boundFlag = TestIncludeTotalBinding -IncludeTotalCount

if ($boundFlag -ne $true) {
    Write-Host "FAIL: IncludeTotalCount was not true when supplied"
    exit 1
}

Write-Host "PASS"
exit 0
