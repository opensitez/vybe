# vybe-test: powershell/csv_delimiter_and_quoting/csv_roundtrip_with_delimiter_preservation
$orig = @(
    [pscustomobject]@{ First = "A"; Second = "B" }
)
$csv = $orig | ConvertTo-Csv -Delimiter '|' -NoTypeInformation
$recovered = $csv | ConvertFrom-Csv -Delimiter '|'
if ($recovered.First -ne "A" -or $recovered.Second -ne "B") {
    Write-Host "FAIL: Custom delimiter roundtrip failed"
    exit 1
}
Write-Host "PASS"
exit 0
