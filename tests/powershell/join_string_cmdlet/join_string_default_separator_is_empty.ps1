# vybe-test: powershell/join_string_cmdlet/join_string_default_separator_is_empty
# When no -Separator is specified, Join-String concatenates items without delimiter
$res = @('a', 'b', 'c') | Join-String

if ($null -eq $res) {
    Write-Host "FAIL: Join-String returned `$null"
    exit 1
}

if ($res -ne "abc") {
    Write-Host "FAIL: expected 'abc', got: '$res'"
    exit 1
}

Write-Host "PASS"
exit 0
