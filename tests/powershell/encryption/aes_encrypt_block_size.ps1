# vybe-test: powershell/encryption/aes_encrypt_block_size
$aes = [System.Security.Cryptography.Aes]::Create()
if ($aes.BlockSize -lt 128) {
    Write-Host "FAIL: expected block size"
    exit 1
}
Write-Host 'PASS'
exit 0
