# vybe-test: powershell/supports_paging_parameters/supports_paging_first_default_value_is_max_uint64
# When the caller omits -First, $PSCmdlet.PagingParameters.First defaults to [ulong]::MaxValue
function TestFirstDefault {
    [CmdletBinding(SupportsPaging = $true)]
    param()
    process {
        return $PSCmdlet.PagingParameters.First
    }
}

$defaultFirst = TestFirstDefault

if ($defaultFirst -ne [System.UInt64]::MaxValue) {
    Write-Host "FAIL: expected First default UInt64.MaxValue, got: $defaultFirst"
    exit 1
}

Write-Host "PASS"
exit 0
