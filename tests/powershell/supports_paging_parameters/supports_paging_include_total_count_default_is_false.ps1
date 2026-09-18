# vybe-test: powershell/supports_paging_parameters/supports_paging_include_total_count_default_is_false
# When -IncludeTotalCount is not specified, $PSCmdlet.PagingParameters.IncludeTotalCount evaluates to false
function TestIncludeTotalCountDefault {
    [CmdletBinding(SupportsPaging = $true)]
    param()
    process {
        return $PSCmdlet.PagingParameters.IncludeTotalCount.IsPresent
    }
}

$defaultIncludeTotal = TestIncludeTotalCountDefault

if ($defaultIncludeTotal) {
    Write-Host "FAIL: IncludeTotalCount was unexpectedly true by default"
    exit 1
}

Write-Host "PASS"
exit 0
