# vybe-test: powershell/supports_paging_parameters/supports_paging_skip_supplied_value_binding
# Passing -Skip binds the specified unsigned integer to $PSCmdlet.PagingParameters.Skip
function TestSkipBinding {
    [CmdletBinding(SupportsPaging = $true)]
    param()
    process {
        return $PSCmdlet.PagingParameters.Skip
    }
}

$boundSkip = TestSkipBinding -Skip 120

if ($boundSkip -ne 120UL) {
    Write-Host "FAIL: expected Skip 120, got: $boundSkip"
    exit 1
}

Write-Host "PASS"
exit 0
