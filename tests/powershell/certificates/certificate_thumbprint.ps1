# vybe-test: powershell/certificates/certificate_thumbprint
$cert = New-Object System.Security.Cryptography.X509Certificates.X509Certificate2
if ($cert.Thumbprint -eq $null) {
    Write-Host "FAIL: expected thumbprint"
    exit 1
}
Write-Host 'PASS'
exit 0
