# vybe-test: powershell/type_casting/string_to_bool
$value = [bool]'True'
if ($value -ne $true) {
    Write-Host 'FAIL'
    exit 1
}
Write-Host 'PASS'
exit 0
