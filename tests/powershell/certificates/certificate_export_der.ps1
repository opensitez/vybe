# vybe-test: powershell/certificates/certificate_export_der
$rsa = [System.Security.Cryptography.RSA]::Create(2048)
$req = [System.Security.Cryptography.X509Certificates.CertificateRequest]::new(
    "CN=ExportTest",
    $rsa,
    [System.Security.Cryptography.HashAlgorithmName]::SHA256,
    [System.Security.Cryptography.RSASignaturePadding]::Pkcs1
)
$cert = $req.CreateSelfSigned([datetimeoffset]::UtcNow, [datetimeoffset]::UtcNow.AddDays(1))
$rawBytes = $cert.Export([System.Security.Cryptography.X509Certificates.X509ContentType]::Cert)
if ($rawBytes.Length -eq 0) {
    Write-Host "FAIL: Certificate Export DER bytes empty"
    exit 1
}
Write-Host "PASS"
exit 0
