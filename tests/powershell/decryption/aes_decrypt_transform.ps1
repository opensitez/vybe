# vybe-test: powershell/decryption/aes_decrypt_transform
$aes = [System.Security.Cryptography.Aes]::Create()
$enc = $aes.CreateEncryptor()
$dec = $aes.CreateDecryptor()
$bytes = $enc.TransformFinalBlock([System.Text.Encoding]::UTF8.GetBytes('b'), 0, 1)
$result = $dec.TransformFinalBlock($bytes, 0, $bytes.Length)
if ([System.Text.Encoding]::UTF8.GetString($result) -ne 'b') {
    Write-Host "FAIL"
    exit 1
}
Write-Host 'PASS'
exit 0
