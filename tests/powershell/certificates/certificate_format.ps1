# vybe-test: powershell/certificates/certificate_format
$rsa = [System.Security.Cryptography.RSA]::Create(2048)
$req = [System.Security.Cryptography.X509Certificates.CertificateRequest]::new(
    "CN=FormatTest",
    $rsa,
    [System.Security.Cryptography.HashAlgorithmName]::SHA256,
    [System.Security.Cryptography.RSASignaturePadding]::Pkcs1
)
$cert = $req.CreateSelfSigned([datetimeoffset]::UtcNow, [datetimeoffset]::UtcNow.AddDays(1))
$str = $cert.ToString($true)
if (-not $str.Contains("CN=FormatTest")) {
    Write-Host "FAIL: Certificate ToString format failed"
    exit 1
}
Write-Host "PASS"
exit 0
