# vybe-test: powershell/format_wide_engine/wide_pipeline_streaming_into_out_string
$records = 1..6 | ForEach-Object { [pscustomobject]@{ Val = "Val$_" } }

# Piping Format-Wide into Out-String produces multi-column formatted text
$strOutput = $records | Format-Wide -Property Val -Column 2 | Out-String

if ($null -eq $strOutput) {
    Write-Host "FAIL: output was null"
    exit 1
}

if (-not ($strOutput -match "Val1" -and $strOutput -match "Val6")) {
    Write-Host "FAIL: streaming into Out-String missing items: $strOutput"
    exit 1
}

Write-Host "PASS"
exit 0
