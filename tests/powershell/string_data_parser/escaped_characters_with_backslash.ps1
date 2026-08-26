# vybe-test: powershell/string_data_parser/escaped_characters_with_backslash
$str = "path = C:\\temp\\file.txt`nmsg = line1\nline2"
$ht = ConvertFrom-StringData -StringData $str
if (-not $ht["path"].Contains("\") -or -not $ht["msg"].Contains("`n")) {
    Write-Host "FAIL: Backslash escapes handling failed"
    exit 1
}
Write-Host "PASS"
exit 0
