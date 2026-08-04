# vybe-test: powershell/type_casting/string_to_guid
$value = [guid]'00000000-0000-0000-0000-000000000000'
if ($value -eq $null) {
    Write-Host 'FAIL'
    exit 1
}
Write-Host 'PASS'
exit 0
