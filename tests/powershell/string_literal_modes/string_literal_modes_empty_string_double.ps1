# vybe-test: powershell/string_literal_modes/empty_string_double
$empty = ""
if ($empty.Length -ne 0) {
    Write-Host "FAIL: double-quoted empty string length should be 0, got $($empty.Length)"
    exit 1
}

Write-Host 'PASS'
exit 0
