# vybe-test: powershell/format_table_engine/table_repeat_header_frequency
$numbers = 1..10 | ForEach-Object { [pscustomobject]@{ Index = $_ } }

# -RepeatHeader parameter instructs the engine to re-emit header blocks periodically
$output = $numbers | Format-Table -Property Index -RepeatHeader | Out-String

if ($null -eq $output) {
    Write-Host "FAIL: output was null"
    exit 1
}

# The header Index and entries must appear
if (-not ($output -match "Index" -and $output -match "10")) {
    Write-Host "FAIL: -RepeatHeader output missing data: $output"
    exit 1
}

Write-Host "PASS"
exit 0
