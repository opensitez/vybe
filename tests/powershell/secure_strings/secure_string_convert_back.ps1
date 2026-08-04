# vybe-test: powershell/secure_strings/secure_string_convert_back
$secure = ConvertTo-SecureString 'secret' -AsPlainText -Force
$plain = [Runtime.InteropServices.Marshal]::PtrToStringAuto([Runtime.InteropServices.Marshal]::SecureStringToBSTR($secure))
if ($plain -ne 'secret') {
    Write-Host "FAIL: expected secret"
    exit 1
}
Write-Host 'PASS'
exit 0
