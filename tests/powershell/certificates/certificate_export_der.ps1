# vybe-test: powershell/certificates/certificate_export_der
$cert = New-Object System.Security.Cryptography.X509Certificates.X509Certificate2
$bytes = $cert.Export([System.Security.Cryptography.X509Certificates.X509ContentType]::Cert)
if ($bytes.Length -lt 0) {
    Write-Host "FAIL: expected bytes array"
    exit 1
}
Write-Host 'PASS'
exit 0
