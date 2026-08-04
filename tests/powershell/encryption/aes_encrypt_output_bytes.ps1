# vybe-test: powershell/encryption/aes_encrypt_output_bytes
$aes = [System.Security.Cryptography.Aes]::Create()
$encryptor = $aes.CreateEncryptor()
$encrypted = $encryptor.TransformFinalBlock([System.Text.Encoding]::UTF8.GetBytes('hello'), 0, 5)
if ($encrypted.Length -le 0) {
    Write-Host "FAIL: expected bytes"
    exit 1
}
Write-Host 'PASS'
exit 0
