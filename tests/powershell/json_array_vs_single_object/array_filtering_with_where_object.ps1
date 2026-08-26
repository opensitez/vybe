# vybe-test: powershell/json_array_vs_single_object/array_filtering_with_where_object
$json = '[{"name":"A","val":1},{"name":"B","val":2},{"name":"C","val":3}]'
$filtered = @(ConvertFrom-Json -InputObject $json | Where-Object { $_.val -gt 1 })
if ($filtered.Length -ne 2 -or $filtered[0].name -ne "B" -or $filtered[1].name -ne "C") {
    Write-Host "FAIL: Where-Object on parsed JSON array failed"
    exit 1
}
Write-Host "PASS"
exit 0
