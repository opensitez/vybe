# vybe-test: powershell/secure_strings/secure_string_not_plain
$secure = ConvertTo-SecureString 'x' -AsPlainText -Force
if ($secure.ToString() -like '*x*') {
    Write-Host "FAIL: secure string should not reveal plain text"
    exit 1
}
Write-Host 'PASS'
exit 0
