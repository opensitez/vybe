# vybe-test: powershell/encryption/aes_encrypt_iv_length
$aes = [System.Security.Cryptography.Aes]::Create()
if ($aes.IV.Length -le 0) {
    Write-Host "FAIL: expected IV length"
    exit 1
}
Write-Host 'PASS'
exit 0
