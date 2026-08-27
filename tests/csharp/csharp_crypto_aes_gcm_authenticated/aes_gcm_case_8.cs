// vybe-test: csharp/csharp_crypto_aes_gcm_authenticated/aes_gcm_case_8

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

bool supported = System.Security.Cryptography.AesGcm.IsSupported;
__P(supported.ToString());
__Check("True");
