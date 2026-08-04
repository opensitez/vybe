# vybe-test: powershell/certificates/certificate_verify
$cert = New-Object System.Security.Cryptography.X509Certificates.X509Certificate2
$verified = $cert.Verify()
if ($verified -ne $false -and $verified -ne $true) {
    Write-Host "FAIL: expected boolean verify"
    exit 1
}
Write-Host 'PASS'
exit 0
