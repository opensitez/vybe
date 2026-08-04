# vybe-test: powershell/encryption/aes_encrypt_key_length
$aes = [System.Security.Cryptography.Aes]::Create()
if ($aes.Key.Length -le 0) {
    Write-Host "FAIL: expected key length"
    exit 1
}
Write-Host 'PASS'
exit 0
