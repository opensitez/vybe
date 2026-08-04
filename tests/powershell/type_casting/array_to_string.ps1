# vybe-test: powershell/type_casting/array_to_string
$value = [string](1,2,3)
if ($value -ne '1 2 3') {
    Write-Host 'FAIL'
    exit 1
}
Write-Host 'PASS'
exit 0
