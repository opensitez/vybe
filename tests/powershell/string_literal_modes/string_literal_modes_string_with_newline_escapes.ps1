# vybe-test: powershell/string_literal_modes/string_with_newline_escapes
$value = "a`nb"
if ($value -ne "a`nb") {
    Write-Host 'FAIL: escaped newline expected in evaluated double-quoted string'
    exit 1
}

if ($value.Contains("`n") -ne $true) {
    Write-Host 'FAIL: escaped newline was not interpreted'
    exit 1
}

Write-Host 'PASS'
exit 0
