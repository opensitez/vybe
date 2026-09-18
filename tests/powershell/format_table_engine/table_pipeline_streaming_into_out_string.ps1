# vybe-test: powershell/format_table_engine/table_pipeline_streaming_into_out_string
$records = 1..3 | ForEach-Object { [pscustomobject]@{ ItemId = $_; Suffix = "s$_" } }

# Streaming Format-Table into Out-String produces multi-line formatted text with header dividers
$strOutput = $records | Format-Table -Property ItemId, Suffix | Out-String

if ($null -eq $strOutput) {
    Write-Host "FAIL: output was null"
    exit 1
}

# The output must contain lines for headers, dividers, and all 3 data rows
$lines = $strOutput.Trim().Split("`n") | Where-Object { $_.Trim().Length -gt 0 }

# At least Header, Divider, and 3 rows = 5 lines
if ($lines.Count -lt 5) {
    Write-Host "FAIL: expected at least 5 rendered lines in table string output, got: $($lines.Count)"
    exit 1
}

if (-not ($strOutput -match "s1" -and $strOutput -match "s2" -and $strOutput -match "s3")) {
    Write-Host "FAIL: data records missing from formatted string output: $strOutput"
    exit 1
}

Write-Host "PASS"
exit 0
