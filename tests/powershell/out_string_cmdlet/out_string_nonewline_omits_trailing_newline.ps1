# vybe-test: powershell/out_string_cmdlet/out_string_nonewline_omits_trailing_newline
# The -NoNewline parameter strips the trailing carriage return and line feed
$output = "pure_content" | Out-String -NoNewline

if ($output.Length -ne 12) {
    Write-Host "FAIL: string length mismatch with -NoNewline, expected 12, got: $($output.Length)"
    exit 1
}

if ($output -ne "pure_content") {
    Write-Host "FAIL: content mismatch, got: '$output'"
    exit 1
}

Write-Host "PASS"
exit 0
