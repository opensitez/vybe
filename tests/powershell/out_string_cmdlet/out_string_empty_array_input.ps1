# vybe-test: powershell/out_string_cmdlet/out_string_empty_array_input
# Piping an empty array into Out-String produces an empty string without errors
$output = @() | Out-String

if ($output -ne "") {
    Write-Host "FAIL: expected empty string from empty array, got: '$output'"
    exit 1
}

Write-Host "PASS"
exit 0
