# vybe-test: powershell/certificates/certificate_import_from_file
$rsa = [System.Security.Cryptography.RSA]::Create(2048)
$req = [System.Security.Cryptography.X509Certificates.CertificateRequest]::new(
    "CN=ImportCertTest",
    $rsa,
    [System.Security.Cryptography.HashAlgorithmName]::SHA256,
    [System.Security.Cryptography.RSASignaturePadding]::Pkcs1
)
$cert = $req.CreateSelfSigned([datetimeoffset]::UtcNow, [datetimeoffset]::UtcNow.AddDays(1))
$exported = $cert.Export([System.Security.Cryptography.X509Certificates.X509ContentType]::Cert)
$imported = [System.Security.Cryptography.X509Certificates.X509Certificate2]::new($exported)
if ($imported.Subject -ne "CN=ImportCertTest") {
    Write-Host "FAIL: Certificate export/import roundtrip failed"
    exit 1
}
Write-Host "PASS"
exit 0
