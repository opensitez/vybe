# vybe-test: powershell/formatting/format_operator_numeric
$pi = 3.14159265
$result = "{0:F2}" -f $pi
if ($result -ne "3.14") {
    Write-Host "FAIL: expected '3.14', got '$result'"
    exit 1
}
$big = "{0:N0}" -f 1234567
if ($big -ne "1,234,567") {
    Write-Host "FAIL: expected '1,234,567', got '$big'"
    exit 1
}
Write-Host "PASS"
exit 0
