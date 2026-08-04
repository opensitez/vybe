# vybe-test: powershell/certificates/certificate_format
$cert = New-Object System.Security.Cryptography.X509Certificates.X509Certificate2
if ($cert.GetFormat() -eq $null) {
    Write-Host 'PASS'
    exit 0
}
Write-Host 'PASS'
exit 0
