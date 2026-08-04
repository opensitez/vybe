# vybe-test: powershell/certificates/create_certificate_object
$cert = New-Object System.Security.Cryptography.X509Certificates.X509Certificate2
if (-not $cert) {
    Write-Host "FAIL: expected certificate object"
    exit 1
}
Write-Host 'PASS'
exit 0
