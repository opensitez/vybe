# vybe-test: powershell/json_array_vs_single_object/array_sorting_with_sort_object
$json = '[{"n":3},{"n":1},{"n":2}]'
$sorted = @(ConvertFrom-Json -InputObject $json | Sort-Object -Property n)
if ($sorted[0].n -ne 1 -or $sorted[2].n -ne 3) {
    Write-Host "FAIL: Sort-Object on parsed JSON array failed"
    exit 1
}
Write-Host "PASS"
exit 0
