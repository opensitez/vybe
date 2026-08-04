# vybe-test: powershell/indexing/substring_indexing
$str = 'PowerShell'
if ($str.Substring(0,5) -ne 'Power') {
    Write-Host 'FAIL'
    exit 1
}
Write-Host 'PASS'
exit 0
