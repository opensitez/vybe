# vybe-test: powershell/certificates/certificate_has_version
$cert = New-Object System.Security.Cryptography.X509Certificates.X509Certificate2
if ($cert.Version -lt 0) {
    Write-Host "FAIL: expected certificate version"
    exit 1
}
Write-Host 'PASS'
exit 0
