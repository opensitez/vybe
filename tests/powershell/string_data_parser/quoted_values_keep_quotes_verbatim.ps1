# vybe-test: powershell/string_data_parser/quoted_values_keep_quotes_verbatim
$str = 'quoted = "hello world"'
$ht = ConvertFrom-StringData -StringData $str
if ($ht["quoted"] -ne '"hello world"') {
    Write-Host "FAIL: Quoted value verbatim preservation failed, got '$($ht["quoted"])'"
    exit 1
}
Write-Host "PASS"
exit 0
