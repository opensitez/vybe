# vybe-test: powershell/supports_paging_parameters/supports_paging_skip_default_value_is_zero
# When the caller omits -Skip, $PSCmdlet.PagingParameters.Skip defaults to 0
function TestSkipDefault {
    [CmdletBinding(SupportsPaging = $true)]
    param()
    process {
        return $PSCmdlet.PagingParameters.Skip
    }
}

$defaultSkip = TestSkipDefault

if ($defaultSkip -ne 0UL) {
    Write-Host "FAIL: expected Skip default 0, got: $defaultSkip"
    exit 1
}

Write-Host "PASS"
exit 0
