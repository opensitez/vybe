# vybe-test: powershell/string_literal_modes/here_strings_no_trim
$literal = @'
C:\Temp\$value
'@
if ($literal -notlike '*$value*') {
    Write-Host "FAIL: single-quoted here-string should keep literal $value token"
    exit 1
}

Write-Host 'PASS'
exit 0
