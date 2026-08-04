# vybe-test: powershell/decryption/aes_decrypt_roundtrip_text
$aes = [System.Security.Cryptography.Aes]::Create()
$key = $aes.Key
$iv = $aes.IV
$en = $aes.CreateEncryptor($key, $iv).TransformFinalBlock([System.Text.Encoding]::UTF8.GetBytes('PS'),0,2)
$de = $aes.CreateDecryptor($key, $iv).TransformFinalBlock($en,0,$en.Length)
if ([System.Text.Encoding]::UTF8.GetString($de) -ne 'PS') {
    Write-Host "FAIL"
    exit 1
}
Write-Host 'PASS'
exit 0
