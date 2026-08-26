# vybe-test: powershell/csv_header_manipulation/header_inspection_via_psobject_properties
$csv = @"
ColA,ColB,ColC
1,2,3
"@
$row = $csv | ConvertFrom-Csv
$headers = @($row.PSObject.Properties | ForEach-Object { $_.Name })
if ($headers.Length -ne 3 -or $headers[0] -ne "ColA" -or $headers[2] -ne "ColC") {
    Write-Host "FAIL: Header inspection via PSObject.Properties failed"
    exit 1
}
Write-Host "PASS"
exit 0
