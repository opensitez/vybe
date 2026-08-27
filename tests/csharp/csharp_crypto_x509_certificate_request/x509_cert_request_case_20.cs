// vybe-test: csharp/csharp_crypto_x509_certificate_request/x509_cert_request_case_20

string __buf = "";

void __P(string s) {
    __buf = __buf + s + "\n";
}

void __Pr(string s) {
    __buf = __buf + s;
}

void __Check(string want) {
    if (__buf != want && __buf != want + "\n") {
        Console.WriteLine("FAIL: want [" + want + "] got [" + __buf + "]");
        throw new Exception("assertion failed");
    }
}

using var rsa = System.Security.Cryptography.RSA.Create(2048);
var req = new System.Security.Cryptography.X509Certificates.CertificateRequest("CN=TestCert_20", rsa, System.Security.Cryptography.HashAlgorithmName.SHA256, System.Security.Cryptography.RSASignaturePadding.Pkcs1);
using var cert = req.CreateSelfSigned(DateTimeOffset.UtcNow, DateTimeOffset.UtcNow.AddDays(1));
__P(cert.Subject);
__Check("CN=TestCert_20");
