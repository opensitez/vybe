// vybe-test: csharp/csharp_crypto_hash/hex_encoding_of_sha256_is_64_chars_long
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

var hash=System.Security.Cryptography.SHA256.HashData(System.Text.Encoding.UTF8.GetBytes("test"));
string hex=System.Convert.ToHexString(hash);
__P((hex.Length).ToString());
__Check("64");
