# vybe-test: powershell/out_string_cmdlet/out_string_null_input
# Piping $null into Out-String evaluates to an empty string
$output = $null | Out-String

if ($output -ne "") {
    Write-Host "FAIL: expected empty string from `$null input, got: '$output'"
    exit 1
}

Write-Host "PASS"
exit 0
