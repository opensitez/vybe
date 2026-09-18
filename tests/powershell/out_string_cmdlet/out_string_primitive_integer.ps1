# vybe-test: powershell/out_string_cmdlet/out_string_primitive_integer
# Primitive integers convert cleanly to their string representation with Out-String -NoNewline
$output = 4096 | Out-String -NoNewline

if ($output -ne "4096") {
    Write-Host "FAIL: expected '4096', got: '$output'"
    exit 1
}

Write-Host "PASS"
exit 0
