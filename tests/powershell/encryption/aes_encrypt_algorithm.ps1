# vybe-test: powershell/encryption/aes_encrypt_algorithm
$aes = [System.Security.Cryptography.Aes]::Create()
if ($aes.GetType().Name -notlike '*Aes*') {
    Write-Host "FAIL: expected Aes algorithm"
    exit 1
}
Write-Host 'PASS'
exit 0
