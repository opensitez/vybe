# vybe-test: powershell/certificates/certificate_verify
$rsa = [System.Security.Cryptography.RSA]::Create(2048)
$req = [System.Security.Cryptography.X509Certificates.CertificateRequest]::new(
    "CN=VerifyTest",
    $rsa,
    [System.Security.Cryptography.HashAlgorithmName]::SHA256,
    [System.Security.Cryptography.RSASignaturePadding]::Pkcs1
)
$cert = $req.CreateSelfSigned([datetimeoffset]::UtcNow, [datetimeoffset]::UtcNow.AddDays(1))
$notBefore = $cert.NotBefore
$notAfter = $cert.NotAfter
if ($notAfter -le $notBefore) {
    Write-Host "FAIL: Certificate validity period check failed"
    exit 1
}
Write-Host "PASS"
exit 0
