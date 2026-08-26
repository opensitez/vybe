# vybe-test: powershell/certificates/certificate_has_version
$rsa = [System.Security.Cryptography.RSA]::Create(2048)
$req = [System.Security.Cryptography.X509Certificates.CertificateRequest]::new(
    "CN=VerTest",
    $rsa,
    [System.Security.Cryptography.HashAlgorithmName]::SHA256,
    [System.Security.Cryptography.RSASignaturePadding]::Pkcs1
)
$cert = $req.CreateSelfSigned([datetimeoffset]::UtcNow, [datetimeoffset]::UtcNow.AddDays(1))
if ($cert.Version -ne 3) {
    Write-Host "FAIL: Certificate Version expected 3, got $($cert.Version)"
    exit 1
}
Write-Host "PASS"
exit 0
