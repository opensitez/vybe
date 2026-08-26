# vybe-test: powershell/csv_delimiter_and_quoting/converto_csv_with_custom_delimiter
$items = @(
    [pscustomobject]@{ K = 1; V = "A" },
    [pscustomobject]@{ K = 2; V = "B" }
)
$csvLines = @($items | ConvertTo-Csv -Delimiter ';' -NoTypeInformation)
if (-not $csvLines[0].Contains("K;V") -and -not $csvLines[0].Contains('"K";"V"')) {
    Write-Host "FAIL: ConvertTo-Csv custom delimiter failed, got '$($csvLines[0])'"
    exit 1
}
Write-Host "PASS"
exit 0
