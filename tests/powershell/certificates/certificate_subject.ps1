# vybe-test: powershell/certificates/certificate_subject
$rsa = [System.Security.Cryptography.RSA]::Create(2048)
$req = [System.Security.Cryptography.X509Certificates.CertificateRequest]::new(
    "CN=VybeServer",
    $rsa,
    [System.Security.Cryptography.HashAlgorithmName]::SHA256,
    [System.Security.Cryptography.RSASignaturePadding]::Pkcs1
)
$cert = $req.CreateSelfSigned([datetimeoffset]::UtcNow, [datetimeoffset]::UtcNow.AddDays(1))
if (-not $cert.Subject.Contains("VybeServer")) {
    Write-Host "FAIL: Certificate Subject check failed"
    exit 1
}
Write-Host "PASS"
exit 0
