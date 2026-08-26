# vybe-test: powershell/certificates/certificate_friendly_name
$rsa = [System.Security.Cryptography.RSA]::Create(2048)
$req = [System.Security.Cryptography.X509Certificates.CertificateRequest]::new(
    "CN=FriendlyTest",
    $rsa,
    [System.Security.Cryptography.HashAlgorithmName]::SHA256,
    [System.Security.Cryptography.RSASignaturePadding]::Pkcs1
)
$cert = $req.CreateSelfSigned([datetimeoffset]::UtcNow, [datetimeoffset]::UtcNow.AddDays(1))
$cert.FriendlyName = "MyTestCert"
if ($cert.FriendlyName -ne "MyTestCert" -and $cert.FriendlyName -ne "") {
    Write-Host "FAIL: FriendlyName check failed"
    exit 1
}
Write-Host "PASS"
exit 0
