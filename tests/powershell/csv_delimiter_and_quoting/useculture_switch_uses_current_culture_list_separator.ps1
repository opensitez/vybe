# vybe-test: powershell/csv_delimiter_and_quoting/useculture_switch_uses_current_culture_list_separator
$sep = [System.Globalization.CultureInfo]::CurrentCulture.TextInfo.ListSeparator
$items = @([pscustomobject]@{ A = 1; B = 2 })
$csvLines = @($items | ConvertTo-Csv -UseCulture -NoTypeInformation)
if (-not $csvLines[0].Contains($sep)) {
    Write-Host "FAIL: ConvertTo-Csv -UseCulture failed"
    exit 1
}
Write-Host "PASS"
exit 0
