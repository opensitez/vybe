# vybe-test: powershell/decryption/aes_decrypt_empty
$aes = [System.Security.Cryptography.Aes]::Create()
$enc = $aes.CreateEncryptor()
$data = [System.Text.Encoding]::UTF8.GetBytes('a')
$encrypted = $enc.TransformFinalBlock($data, 0, $data.Length)
$dec = $aes.CreateDecryptor()
$decrypted = $dec.TransformFinalBlock($encrypted, 0, $encrypted.Length)
if ([System.Text.Encoding]::UTF8.GetString($decrypted) -ne 'a') {
    Write-Host "FAIL: expected a"
    exit 1
}
Write-Host 'PASS'
exit 0
