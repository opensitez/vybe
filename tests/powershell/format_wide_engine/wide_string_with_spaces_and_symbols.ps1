# vybe-test: powershell/format_wide_engine/wide_string_with_spaces_and_symbols
$items = @(
    [pscustomobject]@{ Description = "First Item (Code #101)" },
    [pscustomobject]@{ Description = "Second Item [Status: OK]" }
)

# Strings containing punctuation, spaces, and brackets format cleanly in wide columns
$output = $items | Format-Wide -Property Description -Column 2 | Out-String

if ($null -eq $output) {
    Write-Host "FAIL: output was null"
    exit 1
}

if (-not ($output -match "First Item \(Code #101\)" -and $output -match "Second Item \[Status: OK\]")) {
    Write-Host "FAIL: strings with spaces and symbols failed to render: $output"
    exit 1
}

Write-Host "PASS"
exit 0
