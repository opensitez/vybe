// vybe-test: csharp/csharp_crypto_sha3_256_sha3_512/sha3_availability_case_6

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

bool supported = System.Security.Cryptography.SHA3_256.IsSupported;
__P((supported == true || supported == false).ToString());
__Check("True");
