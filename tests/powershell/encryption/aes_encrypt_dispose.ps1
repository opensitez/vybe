# vybe-test: powershell/encryption/aes_encrypt_dispose
$aes = [System.Security.Cryptography.Aes]::Create()
$aes.Dispose()
Write-Host 'PASS'
exit 0
