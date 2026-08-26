# vybe-test: powershell/certificates/certificate_raw_data
$rsa = [System.Security.Cryptography.RSA]::Create(2048)
$req = [System.Security.Cryptography.X509Certificates.CertificateRequest]::new(
    "CN=RawTest",
    $rsa,
    [System.Security.Cryptography.HashAlgorithmName]::SHA256,
    [System.Security.Cryptography.RSASignaturePadding]::Pkcs1
)
$cert = $req.CreateSelfSigned([datetimeoffset]::UtcNow, [datetimeoffset]::UtcNow.AddDays(1))
$data = $cert.RawData
if ($data.Length -le 0) {
    Write-Host "FAIL: RawData property check failed"
    exit 1
}
Write-Host "PASS"
exit 0
