# vybe-test: powershell/csv_delimiter_and_quoting/converto_csv_notypeinformation_switch
$items = @([pscustomobject]@{ Name = "Test" })
$csvLines = @($items | ConvertTo-Csv -NoTypeInformation)
if ($csvLines[0].StartsWith("#TYPE")) {
    Write-Host "FAIL: -NoTypeInformation should omit #TYPE header line"
    exit 1
}
Write-Host "PASS"
exit 0
