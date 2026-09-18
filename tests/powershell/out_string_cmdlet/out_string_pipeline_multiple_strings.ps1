# vybe-test: powershell/out_string_cmdlet/out_string_pipeline_multiple_strings
# Streaming multiple string elements through Out-String combines them into newline-separated text
$elements = @("first_row", "second_row", "third_row")
$output = $elements | Out-String

if ($output -notmatch "first_row" -or $output -notmatch "second_row" -or $output -notmatch "third_row") {
    Write-Host "FAIL: elements missing from aggregated Out-String text"
    exit 1
}

Write-Host "PASS"
exit 0
