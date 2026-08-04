# vybe-test: powershell/certificates/certificate_subject
$cert = New-Object System.Security.Cryptography.X509Certificates.X509Certificate2
if ($cert.Subject -eq $null) {
    Write-Host "FAIL: expected certificate subject"
    exit 1
}
Write-Host 'PASS'
exit 0
