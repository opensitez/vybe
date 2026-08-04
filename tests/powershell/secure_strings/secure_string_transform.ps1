# vybe-test: powershell/secure_strings/secure_string_transform
$secure = ConvertTo-SecureString 'z' -AsPlainText -Force
if (-not $secure.Length) {
    Write-Host "FAIL: expected secure string length"
    exit 1
}
Write-Host 'PASS'
exit 0
