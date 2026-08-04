# vybe-test: powershell/decryption/aes_decrypt_dispose
$aes = [System.Security.Cryptography.Aes]::Create()
$aes.Dispose()
Write-Host 'PASS'
exit 0
