# vybe-test: powershell/join_string_cmdlet/join_string_custom_separator
# -Separator specifies the delimiter between joined items
$res = @('red', 'green', 'blue') | Join-String -Separator ', '

if ($null -eq $res) {
    Write-Host "FAIL: output was null"
    exit 1
}

if ($res -ne "red, green, blue") {
    Write-Host "FAIL: expected 'red, green, blue', got: '$res'"
    exit 1
}

Write-Host "PASS"
exit 0
