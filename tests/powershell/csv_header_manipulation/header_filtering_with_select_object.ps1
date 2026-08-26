# vybe-test: powershell/csv_header_manipulation/header_filtering_with_select_object
$csv = @"
A,B,C,D
1,2,3,4
"@
$row = $csv | ConvertFrom-Csv | Select-Object -Property A, C
$props = @($row.PSObject.Properties | ForEach-Object { $_.Name })
if ($props.Length -ne 2 -or $props -contains "B" -or $props -contains "D") {
    Write-Host "FAIL: Selecting subset of CSV header properties failed"
    exit 1
}
Write-Host "PASS"
exit 0
