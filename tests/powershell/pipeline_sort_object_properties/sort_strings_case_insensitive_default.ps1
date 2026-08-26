# vybe-test: powershell/pipeline_sort_object_properties/sort_strings_case_insensitive_default
$words = @("banana", "Apple", "cherry")
$sorted = @($words | Sort-Object)
if ($sorted[0] -ne "Apple" -or $sorted[1] -ne "banana" -or $sorted[2] -ne "cherry") {
    Write-Host "FAIL: Sort-Object case-insensitive default failed"
    exit 1
}
Write-Host "PASS"
exit 0
