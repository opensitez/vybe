# vybe-test: powershell/type_casting/type_check
$value = [int]5
if ($value -isnot [int]) {
    Write-Host 'FAIL'
    exit 1
}
Write-Host 'PASS'
exit 0
