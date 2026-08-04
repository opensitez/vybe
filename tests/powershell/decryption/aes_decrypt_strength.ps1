# vybe-test: powershell/decryption/aes_decrypt_strength
$aes = [System.Security.Cryptography.Aes]::Create()
if ($aes.KeySize -lt 128) {
    Write-Host "FAIL"
    exit 1
}
Write-Host 'PASS'
exit 0
