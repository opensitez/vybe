# vybe-test: powershell/join_string_cmdlet/join_string_integer_ranges
# Joining integer ranges directly converts integers to string representation
$res = (1..5) | Join-String -Separator ''

if ($res -ne "12345") {
    Write-Host "FAIL: expected '12345', got: '$res'"
    exit 1
}

$hyphenated = (1..5) | Join-String -Separator '-'
if ($hyphenated -ne "1-2-3-4-5") {
    Write-Host "FAIL: expected '1-2-3-4-5', got: '$hyphenated'"
    exit 1
}

Write-Host "PASS"
exit 0
