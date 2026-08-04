# vybe-test: powershell/decryption/aes_decrypt_algorithm
$aes = [System.Security.Cryptography.Aes]::Create()
if ($aes.GetType().Name -notlike '*Aes*') {
    Write-Host "FAIL"
    exit 1
}
Write-Host 'PASS'
exit 0
