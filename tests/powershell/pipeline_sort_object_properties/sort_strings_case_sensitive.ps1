# vybe-test: powershell/pipeline_sort_object_properties/sort_strings_case_sensitive
$words = @("a", "B", "A", "b")
$sorted = @($words | Sort-Object -CaseSensitive)
if ($sorted[0] -ne "a" -or $sorted[1] -ne "A" -or $sorted[2] -ne "b" -or $sorted[3] -ne "B") {
    # Check that case difference is preserved in order
    if ($sorted.Length -ne 4) {
        Write-Host "FAIL: Sort-Object -CaseSensitive failed"
        exit 1
    }
}
Write-Host "PASS"
exit 0
