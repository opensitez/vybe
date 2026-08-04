# vybe-test: powershell/encryption/aes_encrypt_padding
$aes = [System.Security.Cryptography.Aes]::Create()
if (-not $aes.Padding) {
    Write-Host "FAIL: expected padding"
    exit 1
}
Write-Host 'PASS'
exit 0
