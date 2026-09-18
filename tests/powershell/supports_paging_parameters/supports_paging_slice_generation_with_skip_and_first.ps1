# vybe-test: powershell/supports_paging_parameters/supports_paging_slice_generation_with_skip_and_first
# Advanced function implementing in-memory paging honors -Skip and -First to return precise sub-ranges
function Get-PagedCollection {
    [CmdletBinding(SupportsPaging = $true)]
    param()
    process {
        $p = $PSCmdlet.PagingParameters
        $all = 1..100

        $skipCount = [int][Math]::Min([uint64]$all.Count, $p.Skip)
        $remaining = $all.Count - $skipCount

        $takeCount = if ($p.First -lt [uint64]$remaining) { [int]$p.First } else { $remaining }

        for ($i = 0; $i -lt $takeCount; $i++) {
            $all[$skipCount + $i]
        }
    }
}

$page = @(Get-PagedCollection -Skip 10 -First 4)

if ($page.Count -ne 4) {
    Write-Host "FAIL: expected 4 items, got $($page.Count)"
    exit 1
}

$expected = @(11, 12, 13, 14)
for ($i = 0; $i -lt 4; $i++) {
    if ($page[$i] -ne $expected[$i]) {
        Write-Host "FAIL: at index ${i}, expected $($expected[$i]), got $($page[$i])"
        exit 1
    }
}

Write-Host "PASS"
exit 0
