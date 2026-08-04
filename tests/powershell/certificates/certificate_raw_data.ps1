# vybe-test: powershell/certificates/certificate_raw_data
$cert = New-Object System.Security.Cryptography.X509Certificates.X509Certificate2
if (-not $cert.RawData) {
    Write-Host "FAIL: expected raw data"
    exit 1
}
Write-Host 'PASS'
exit 0
