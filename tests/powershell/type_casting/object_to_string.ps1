# vybe-test: powershell/type_casting/object_to_string
$value = [string]1
if ($value -ne '1') {
    Write-Host 'FAIL'
    exit 1
}
Write-Host 'PASS'
exit 0
