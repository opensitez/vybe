# vybe-test: powershell/pipeline_sort_object_properties/sort_by_scriptblock_calculated_property
$words = @("elephant", "cat", "hippopotamus", "dog")
$sorted = @($words | Sort-Object { $_.Length })
if ($sorted[0].Length -ne 3 -or $sorted[3].Length -ne 12) {
    Write-Host "FAIL: Sort-Object by scriptblock length failed"
    exit 1
}
Write-Host "PASS"
exit 0
