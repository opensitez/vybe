# vybe-test: powershell/decryption/aes_decrypt_key_vector
$aes = [System.Security.Cryptography.Aes]::Create()
if ($aes.Key.Length -le 0 -or $aes.IV.Length -le 0) {
    Write-Host "FAIL: expected key and iv"
    exit 1
}
Write-Host 'PASS'
exit 0
