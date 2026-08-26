# vybe-test: powershell/pipeline_sort_object_properties/sort_stable_preservation_of_equal_elements
$items = @(
    [pscustomobject]@{ Rank = 1; Tag = "First" },
    [pscustomobject]@{ Rank = 1; Tag = "Second" },
    [pscustomobject]@{ Rank = 1; Tag = "Third" }
)
$sorted = @($items | Sort-Object -Property Rank)
if ($sorted[0].Tag -ne "First" -or $sorted[1].Tag -ne "Second" -or $sorted[2].Tag -ne "Third") {
    Write-Host "FAIL: Sort-Object stable sort preservation failed"
    exit 1
}
Write-Host "PASS"
exit 0
