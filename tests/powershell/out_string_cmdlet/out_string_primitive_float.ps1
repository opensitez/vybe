# vybe-test: powershell/out_string_cmdlet/out_string_primitive_float
# Primitive floating point values convert to string representations
$output = 99.75 | Out-String -NoNewline

if ($output -ne "99.75") {
    Write-Host "FAIL: expected '99.75', got: '$output'"
    exit 1
}

Write-Host "PASS"
exit 0
