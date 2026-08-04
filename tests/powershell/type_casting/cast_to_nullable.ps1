# vybe-test: powershell/type_casting/cast_to_nullable
$value = [int?]5
if ($value -ne 5) {
    Write-Host 'FAIL'
    exit 1
}
Write-Host 'PASS'
exit 0
