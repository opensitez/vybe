# vybe-test: powershell/encryption/aes_encrypt_basic
$aes = [System.Security.Cryptography.Aes]::Create()
$key = $aes.Key
$iv = $aes.IV
$encryptor = $aes.CreateEncryptor($key, $iv)
$data = [System.Text.Encoding]::UTF8.GetBytes('hello')
$encrypted = $encryptor.TransformFinalBlock($data, 0, $data.Length)
if ($encrypted.Length -le 0) {
    Write-Host "FAIL: expected encrypted bytes"
    exit 1
}
Write-Host 'PASS'
exit 0
