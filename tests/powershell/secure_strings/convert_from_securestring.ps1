# vybe-test: powershell/secure_strings/convert_from_securestring
$secure = ConvertTo-SecureString 'password' -AsPlainText -Force
$plain = [Runtime.InteropServices.Marshal]::PtrToStringAuto([Runtime.InteropServices.Marshal]::SecureStringToBSTR($secure))
if ($plain -ne 'password') {
    Write-Host "FAIL: expected password"
    exit 1
}
Write-Host 'PASS'
exit 0
