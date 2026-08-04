# vybe-test: powershell/encryption/aes_encrypt_transform
$aes = [System.Security.Cryptography.Aes]::Create()
$encryptor = $aes.CreateEncryptor()
$data = [System.Text.Encoding]::UTF8.GetBytes('x')
$encrypted = $encryptor.TransformFinalBlock($data, 0, $data.Length)
if ($encrypted.Length -le 0) {
    Write-Host "FAIL: expected bytes"
    exit 1
}
Write-Host 'PASS'
exit 0
