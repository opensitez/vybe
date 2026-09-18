# vybe-test: powershell/supports_paging_parameters/supports_paging_with_pipeline_input
# Advanced functions decorated with SupportsPaging accept and process streamed pipeline input
function FilterPagedPipeline {
    [CmdletBinding(SupportsPaging = $true)]
    param(
        [Parameter(ValueFromPipeline = $true)]
        [int]$InputObject
    )
    begin {
        $index = 0
        $emitted = 0
    }
    process {
        $p = $PSCmdlet.PagingParameters
        if ($index -ge $p.Skip -and ($p.First -eq [System.UInt64]::MaxValue -or $emitted -lt $p.First)) {
            $InputObject * 10
            $emitted++
        }
        $index++
    }
}

$results = @(1..10 | FilterPagedPipeline -Skip 3 -First 2)

# Items at index 3 and 4: 4 and 5 -> multiplied by 10: 40, 50
if ($results.Count -ne 2) {
    Write-Host "FAIL: expected 2 paged pipeline items, got $($results.Count)"
    exit 1
}

if ($results[0] -ne 40 -or $results[1] -ne 50) {
    Write-Host "FAIL: paged pipeline values mismatch: $($results[0]), $($results[1])"
    exit 1
}

Write-Host "PASS"
exit 0
