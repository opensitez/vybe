# vybe-test: powershell/format_table_engine/table_autosize_column_widths
$data = @(
    [pscustomobject]@{ ShortCol = "A"; LongCol = "ThisIsALongerDataEntry" },
    [pscustomobject]@{ ShortCol = "B"; LongCol = "AnotherLongDataEntry" }
)

# -AutoSize calculates column widths based on maximum width of data rows
$output = $data | Format-Table -Property ShortCol, LongCol -AutoSize | Out-String

if ($null -eq $output) {
    Write-Host "FAIL: output was null"
    exit 1
}

# Verify content rendered correctly
if (-not ($output -match "ShortCol" -and $output -match "LongCol" -and $output -match "ThisIsALongerDataEntry")) {
    Write-Host "FAIL: autosized table content missing: $output"
    exit 1
}

Write-Host "PASS"
exit 0
