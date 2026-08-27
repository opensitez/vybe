// vybe-test: csharp/csharp_crypto_rfc2898_pbkdf2_derivation/pbkdf2_case_19

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

byte[] salt = new byte[16];
byte[] key = System.Security.Cryptography.Rfc2898DeriveBytes.Pbkdf2("password", salt, 10, System.Security.Cryptography.HashAlgorithmName.SHA256, 32);
__P(key.Length.ToString());
__Check("32");
