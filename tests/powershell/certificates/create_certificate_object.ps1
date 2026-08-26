# vybe-test: powershell/certificates/create_certificate_object
$rsa = [System.Security.Cryptography.RSA]::Create(2048)
$req = [System.Security.Cryptography.X509Certificates.CertificateRequest]::new(
    "CN=TestCert",
    $rsa,
    [System.Security.Cryptography.HashAlgorithmName]::SHA256,
    [System.Security.Cryptography.RSASignaturePadding]::Pkcs1
)
$cert = $req.CreateSelfSigned([datetimeoffset]::UtcNow, [datetimeoffset]::UtcNow.AddDays(1))
if ($cert.Subject -ne "CN=TestCert") {
    Write-Host "FAIL: Subject expected CN=TestCert, got $($cert.Subject)"
    exit 1
}
Write-Host "PASS"
exit 0
