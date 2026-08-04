# vybe-test: powershell/indexing/string_indexing
$str = 'abc'
if ($str[1] -ne 'b') {
    Write-Host 'FAIL'
    exit 1
}
Write-Host 'PASS'
exit 0
