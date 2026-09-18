# vybe-test: powershell/supports_paging_parameters/supports_paging_zero_first_yields_empty_result
# Supplying -First 0 causes an in-memory paged function to return an empty array
function Get-ZeroFirstSample {
    [CmdletBinding(SupportsPaging = $true)]
    param()
    process {
        $p = $PSCmdlet.PagingParameters
        if ($p.First -eq 0UL) {
            return
        }
        1..10
    }
}

$items = @(Get-ZeroFirstSample -First 0)

if ($items.Count -ne 0) {
    Write-Host "FAIL: expected 0 items when -First 0, got $($items.Count)"
    exit 1
}

Write-Host "PASS"
exit 0
