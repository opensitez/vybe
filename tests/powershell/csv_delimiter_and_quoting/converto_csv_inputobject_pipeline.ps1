# vybe-test: powershell/csv_delimiter_and_quoting/converto_csv_inputobject_pipeline
$csv = 1..3 | ForEach-Object { [pscustomobject]@{ N = $_ } } | ConvertTo-Csv -NoTypeInformation
$rows = @($csv | ConvertFrom-Csv)
if ($rows.Length -ne 3 -or $rows[2].N -ne "3") {
    Write-Host "FAIL: Pipeline ConvertTo-Csv failed"
    exit 1
}
Write-Host "PASS"
exit 0
