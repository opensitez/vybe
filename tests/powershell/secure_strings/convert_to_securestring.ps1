# vybe-test: powershell/secure_strings/convert_to_securestring
$secure = ConvertTo-SecureString 'password' -AsPlainText -Force
if (-not $secure) {
    Write-Host "FAIL: expected secure string"
    exit 1
}
Write-Host 'PASS'
exit 0
