# vybe-test: powershell/string_data_parser/comment_lines_ignored
$str = @"
# This is a comment
name = Alice
# Another comment
role = admin
"@
$ht = ConvertFrom-StringData -StringData $str
if ($ht.Count -ne 2 -or $ht["name"] -ne "Alice" -or $ht["role"] -ne "admin") {
    Write-Host "FAIL: Comment lines not ignored"
    exit 1
}
Write-Host "PASS"
exit 0
