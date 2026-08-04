# vybe-test: powershell/decryption/aes_decrypt_object
$aes = [System.Security.Cryptography.Aes]::Create()
if ($aes -eq $null) {
    Write-Host "FAIL"
    exit 1
}
Write-Host 'PASS'
exit 0
