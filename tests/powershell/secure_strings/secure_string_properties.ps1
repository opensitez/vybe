# vybe-test: powershell/secure_strings/secure_string_properties
$secure = ConvertTo-SecureString 'x' -AsPlainText -Force
if (-not $secure.IsReadOnly()) {
    Write-Host "FAIL: expected secure string property"
    exit 1
}
Write-Host 'PASS'
exit 0
