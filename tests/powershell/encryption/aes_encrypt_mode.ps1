# vybe-test: powershell/encryption/aes_encrypt_mode
$aes = [System.Security.Cryptography.Aes]::Create()
if (-not $aes.Mode) {
    Write-Host "FAIL: expected cipher mode"
    exit 1
}
Write-Host 'PASS'
exit 0
