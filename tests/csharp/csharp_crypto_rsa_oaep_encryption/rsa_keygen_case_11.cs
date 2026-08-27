// vybe-test: csharp/csharp_crypto_rsa_oaep_encryption/rsa_keygen_case_11

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
__P(rsa.KeySize.ToString());
__Check("2048");
