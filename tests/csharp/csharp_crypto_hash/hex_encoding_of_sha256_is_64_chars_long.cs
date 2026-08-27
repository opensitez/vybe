// vybe-test: csharp/csharp_crypto_hash/hex_encoding_of_sha256_is_64_chars_long
// origin: languages/csharp/tests/csharp/test_csharp_crypto_hash.rs

using static __Harness;

var hash=System.Security.Cryptography.SHA256.HashData(System.Text.Encoding.UTF8.GetBytes("test"));
string hex=System.Convert.ToHexString(hash);
__P((hex.Length).ToString());
__Check("64");

public static class __Harness {
    public static string __buf = "";
    public static void __P(string s) { __buf = __buf + s + "\n"; }
    public static void __Pr(string s) { __buf = __buf + s; }
    public static void __Check(string want) {
        if (__buf != want && __buf != want + "\n") {
            Console.WriteLine("FAIL: want [" + want + "] got [" + __buf + "]");
            throw new Exception("assertion failed");
        }
    }
}
