# vybe-test: powershell/pipeline_sort_object_properties/sort_with_null_values_placed_first_in_ascending
$items = @(
    [pscustomobject]@{ Id = 1; Score = 50 },
    [pscustomobject]@{ Id = 2; Score = $null },
    [pscustomobject]@{ Id = 3; Score = 20 }
)
$sorted = @($items | Sort-Object -Property Score)
if ($sorted[0].Id -ne 2 -or $sorted[1].Score -ne 20 -or $sorted[2].Score -ne 50) {
    Write-Host "FAIL: Sort-Object with null values ascending failed"
    exit 1
}
Write-Host "PASS"
exit 0
