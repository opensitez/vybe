# vybe-test: powershell/supports_paging_parameters/supports_paging_first_supplied_value_binding
# Passing -First binds the specified unsigned integer to $PSCmdlet.PagingParameters.First
function TestFirstBinding {
    [CmdletBinding(SupportsPaging = $true)]
    param()
    process {
        return $PSCmdlet.PagingParameters.First
    }
}

$boundFirst = TestFirstBinding -First 50

if ($boundFirst -ne 50UL) {
    Write-Host "FAIL: expected First 50, got: $boundFirst"
    exit 1
}

Write-Host "PASS"
exit 0
