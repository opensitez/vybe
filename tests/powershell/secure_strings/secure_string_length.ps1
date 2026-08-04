# vybe-test: powershell/secure_strings/secure_string_length
$secure = ConvertTo-SecureString 'password' -AsPlainText -Force
if ($secure.Length -lt 1) {
    Write-Host "FAIL: expected secure length"
    exit 1
}
Write-Host 'PASS'
exit 0
