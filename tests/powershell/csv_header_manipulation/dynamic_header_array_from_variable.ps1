# vybe-test: powershell/csv_header_manipulation/dynamic_header_array_from_variable
$headers = @("Col_1", "Col_2")
$csv = "10,20`n30,40"
$rows = @($csv | ConvertFrom-Csv -Header $headers)
if ($rows[0].Col_1 -ne "10" -or $rows[1].Col_2 -ne "40") {
    Write-Host "FAIL: Dynamic header array from variable failed"
    exit 1
}
Write-Host "PASS"
exit 0
