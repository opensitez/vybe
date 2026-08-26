# vybe-test: powershell/csv_header_manipulation/converto_csv_preserves_object_property_order_as_headers
$items = @([pscustomobject][ordered]@{ Z = 1; A = 2; M = 3 })
$csvLines = @($items | ConvertTo-Csv -NoTypeInformation)
if (-not $csvLines[0].Contains('"Z","A","M"') -and -not $csvLines[0].Contains('Z,A,M')) {
    Write-Host "FAIL: Property order preservation in CSV header failed, got '$($csvLines[0])'"
    exit 1
}
Write-Host "PASS"
exit 0
