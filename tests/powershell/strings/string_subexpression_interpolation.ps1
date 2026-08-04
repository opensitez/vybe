# vybe-test: powershell/strings/string_format_interpolation_expression
$a = 6
$b = 7
$result = "The answer is $($a * $b)"
if ($result -ne "The answer is 42") {
    Write-Host "FAIL: got '$result'"
    exit 1
}
Write-Host "PASS"
exit 0
