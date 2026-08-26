# vybe-test: powershell/type_version_parsing_and_comparison/parse_semantic_like_string
$str = "7.3.4"
$v = [version]::Parse($str)
if ($v.Major -ne 7 -or $v.Minor -ne 3 -or $v.Build -ne 4) {
    Write-Host "FAIL: Parse failed for $str"
    exit 1
}
Write-Host "PASS"
exit 0
