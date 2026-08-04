# vybe-test: powershell/secure_strings/secure_string_comparison
$secure = ConvertTo-SecureString 'x' -AsPlainText -Force
$secure2 = ConvertTo-SecureString 'x' -AsPlainText -Force
if ($secure.Length -ne $secure2.Length) {
    Write-Host "FAIL: expected lengths match"
    exit 1
}
Write-Host 'PASS'
exit 0
