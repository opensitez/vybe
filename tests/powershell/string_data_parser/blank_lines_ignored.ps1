# vybe-test: powershell/string_data_parser/blank_lines_ignored
$str = @"

host = localhost

port = 8080

"@
$ht = ConvertFrom-StringData -StringData $str
if ($ht.Count -ne 2 -or $ht["host"] -ne "localhost") {
    Write-Host "FAIL: Blank lines handling failed"
    exit 1
}
Write-Host "PASS"
exit 0
