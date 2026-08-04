# vybe-test: powershell/decryption/aes_decrypt_padding
$aes = [System.Security.Cryptography.Aes]::Create()
if (-not $aes.Padding) {
    Write-Host "FAIL"
    exit 1
}
Write-Host 'PASS'
exit 0
