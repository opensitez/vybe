# vybe-test: powershell/type_casting/array_to_int
$value = [int]('5')
if ($value -ne 5) {
    Write-Host 'FAIL'
    exit 1
}
Write-Host 'PASS'
exit 0
