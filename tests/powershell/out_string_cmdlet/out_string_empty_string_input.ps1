# vybe-test: powershell/out_string_cmdlet/out_string_empty_string_input
# Piping an empty string with -NoNewline preserves the empty string
$output = "" | Out-String -NoNewline

if ($output -ne "") {
    Write-Host "FAIL: expected empty string, got: '$output'"
    exit 1
}

Write-Host "PASS"
exit 0
