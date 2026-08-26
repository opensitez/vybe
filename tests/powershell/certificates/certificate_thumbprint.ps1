# vybe-test: powershell/certificates/certificate_thumbprint
$rsa = [System.Security.Cryptography.RSA]::Create(2048)
$req = [System.Security.Cryptography.X509Certificates.CertificateRequest]::new(
    "CN=ThumbTest",
    $rsa,
    [System.Security.Cryptography.HashAlgorithmName]::SHA256,
    [System.Security.Cryptography.RSASignaturePadding]::Pkcs1
)
$cert = $req.CreateSelfSigned([datetimeoffset]::UtcNow, [datetimeoffset]::UtcNow.AddDays(1))
if ($cert.Thumbprint.Length -ne 40) {
    Write-Host "FAIL: Certificate Thumbprint length expected 40, got $($cert.Thumbprint.Length)"
    exit 1
}
Write-Host "PASS"
exit 0
