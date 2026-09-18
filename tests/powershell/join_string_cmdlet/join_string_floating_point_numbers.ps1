# vybe-test: powershell/join_string_cmdlet/join_string_floating_point_numbers
$floats = @(1.25, 2.5, 3.75)

# Floating point numbers format with decimals preserved
$res = $floats | Join-String -Separator ' | '

if ($res -ne "1.25 | 2.5 | 3.75") {
    Write-Host "FAIL: expected '1.25 | 2.5 | 3.75', got: '$res'"
    exit 1
}

Write-Host "PASS"
exit 0
