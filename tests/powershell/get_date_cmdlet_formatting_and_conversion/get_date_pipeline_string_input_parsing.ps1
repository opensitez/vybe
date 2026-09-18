# vybe-test: powershell/get_date_cmdlet_formatting_and_conversion/get_date_pipeline_string_input_parsing
# Streaming date string representations through the pipeline binds to Get-Date and parses them
$parsed = "2026-07-04" | Get-Date

if ($parsed.Year -ne 2026 -or $parsed.Month -ne 7 -or $parsed.Day -ne 4) {
    Write-Host "FAIL: parsed date mismatch, got: $($parsed.ToString('yyyy-MM-dd'))"
    exit 1
}

Write-Host "PASS"
exit 0
