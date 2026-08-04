// vybe-test: csharp/csharp_crypto_hash/md5_produces_16_byte_digest
// origin: languages/csharp/tests/csharp/test_csharp_crypto_hash.rs

string __buf = "";

void __P(string s) {
    __buf = __buf + s + "\n";
}

void __Pr(string s) {
    __buf = __buf + s;
}

// The final WriteLine contributes a trailing newline that the expected line
// vector never carried, so BOTH forms are accepted.
void __Check(string want) {
    if (__buf != want && __buf != want + "\n") {
        Console.WriteLine("FAIL: want [" + want + "] got [" + __buf + "]");
        throw new Exception("assertion failed");
    }
}

using var md5=System.Security.Cryptography.MD5.Create();
byte[] hash=md5.ComputeHash(new byte[]{1,2,3});
__P((hash.Length).ToString());
__Check("16");
