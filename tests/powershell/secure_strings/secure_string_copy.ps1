# vybe-test: powershell/secure_strings/secure_string_copy
$secure = ConvertTo-SecureString 'y' -AsPlainText -Force
$secure2 = ConvertTo-SecureString 'y' -AsPlainText -Force
if ($secure.Length -ne $secure2.Length) {
    Write-Host "FAIL: expected equal lengths"
    exit 1
}
Write-Host 'PASS'
exit 0
