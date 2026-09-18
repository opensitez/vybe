# vybe-test: powershell/supports_paging_parameters/supports_paging_combined_with_supports_should_process
# SupportsPaging and SupportsShouldProcess coexist cleanly on the same advanced function
function Clear-PagedArchive {
    [CmdletBinding(SupportsPaging = $true, SupportsShouldProcess = $true)]
    param()
    process {
        $p = $PSCmdlet.PagingParameters
        $target = "Page(Skip:$($p.Skip),First:$($p.First))"
        if ($PSCmdlet.ShouldProcess($target, "Purge")) {
            return "Purged:$target"
        }
        return "Skipped:$target"
    }
}

$res = Clear-PagedArchive -Skip 10 -First 5

if ($res -ne "Purged:Page(Skip:10,First:5)") {
    Write-Host "FAIL: combined should_process and paging mismatch: '$res'"
    exit 1
}

Write-Host "PASS"
exit 0
