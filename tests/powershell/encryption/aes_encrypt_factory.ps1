# vybe-test: powershell/encryption/aes_encrypt_factory
$aes = [System.Security.Cryptography.Aes]::Create()
if ($aes -eq $null) {
    Write-Host "FAIL: expected AES instance"
    exit 1
}
Write-Host 'PASS'
exit 0
