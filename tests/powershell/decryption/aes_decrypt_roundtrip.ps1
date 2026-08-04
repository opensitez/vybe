# vybe-test: powershell/decryption/aes_decrypt_roundtrip
$aes = [System.Security.Cryptography.Aes]::Create()
$key = $aes.Key
$iv = $aes.IV
$encryptor = $aes.CreateEncryptor($key, $iv)
$decryptor = $aes.CreateDecryptor($key, $iv)
$data = [System.Text.Encoding]::UTF8.GetBytes('world')
$encrypted = $encryptor.TransformFinalBlock($data, 0, $data.Length)
$decrypted = $decryptor.TransformFinalBlock($encrypted, 0, $encrypted.Length)
$text = [System.Text.Encoding]::UTF8.GetString($decrypted)
if ($text -ne 'world') {
    Write-Host "FAIL: expected world"
    exit 1
}
Write-Host 'PASS'
exit 0
