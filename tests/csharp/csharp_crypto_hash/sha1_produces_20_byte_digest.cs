// vybe-test: csharp/csharp_crypto_hash/sha1_produces_20_byte_digest
// origin: languages/csharp/tests/csharp/test_csharp_crypto_hash.rs

using static __Harness;

var hash=System.Security.Cryptography.SHA1.HashData(new byte[]{0});
__P((hash.Length).ToString());
__Check("20");

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
