# vybe-test: powershell/decryption/aes_decrypt_blocksize.ps1
$aes = [System.Security.Cryptography.Aes]::Create()
if ($aes.BlockSize -lt 128) {
    Write-Host "FAIL"
    exit 1
}
Write-Host 'PASS'
exit 0
