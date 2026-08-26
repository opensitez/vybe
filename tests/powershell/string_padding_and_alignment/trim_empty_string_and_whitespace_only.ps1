# vybe-test: powershell/string_padding_and_alignment/trim_empty_string_and_whitespace_only
$empty = ""
$ws = "   `t`n   "
if ($empty.Trim() -ne "" -or $ws.Trim() -ne "") {
    Write-Host "FAIL: Trim on whitespace/empty string failed"
    exit 1
}
Write-Host "PASS"
exit 0
