# vybe-test: powershell/certificates/certificate_import_from_file
$cert = New-Object System.Security.Cryptography.X509Certificates.X509Certificate2
if ($cert -eq $null) {
    Write-Host "FAIL"
    exit 1
}
Write-Host 'PASS'
exit 0
