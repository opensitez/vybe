# vybe-test: powershell/formatting/format_operator_multiple_args
$result = "{0} has {1} items at ${2:F2} each" -f "Cart", 3, 9.9
if ($result -ne "Cart has 3 items at `$9.90 each") {
    Write-Host "FAIL: got '$result'"
    exit 1
}
Write-Host "PASS"
exit 0
