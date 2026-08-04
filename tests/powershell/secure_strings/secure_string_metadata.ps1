# vybe-test: powershell/secure_strings/secure_string_metadata
$secure = ConvertTo-SecureString 'x' -AsPlainText -Force
if (-not $secure.GetType().Name) {
    Write-Host "FAIL: expected type"
    exit 1
}
Write-Host 'PASS'
exit 0
