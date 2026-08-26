# vybe-test: powershell/string_padding_and_alignment/trim_carriage_return_newline_escapes
$str = "`r`n`tHello`t`r`n"
$trimmed = $str.Trim()
if ($trimmed -ne "Hello") {
    Write-Host "FAIL: Trim CRLF and tabs failed, got '$trimmed'"
    exit 1
}
Write-Host "PASS"
exit 0
