# vybe-test: powershell/supports_paging_parameters/supports_paging_combined_with_custom_parameters
# User-defined parameters function normally alongside automatic paging parameters
function QueryPagedRecords {
    [CmdletBinding(SupportsPaging = $true)]
    param(
        [string]$Category,
        [int]$MinPriority
    )
    process {
        $p = $PSCmdlet.PagingParameters
        return "Cat:$Category,MinP:$MinPriority,First:$($p.First),Skip:$($p.Skip)"
    }
}

$summary = QueryPagedRecords -Category "Security" -MinPriority 3 -First 10 -Skip 5

if ($summary -ne "Cat:Security,MinP:3,First:10,Skip:5") {
    Write-Host "FAIL: combined parameters execution mismatch: '$summary'"
    exit 1
}

Write-Host "PASS"
exit 0
