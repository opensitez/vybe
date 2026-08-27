// vybe-test: csharp/csharp_crypto_hash/sha256_hash_of_same_input_is_always_identical
// origin: languages/csharp/tests/csharp/test_csharp_crypto_hash.rs

using static __Harness;

byte[] input=System.Text.Encoding.UTF8.GetBytes("hello");
var h1=System.Security.Cryptography.SHA256.HashData(input);
var h2=System.Security.Cryptography.SHA256.HashData(input);
__P((System.MemoryExtensions.SequenceEqual<byte>(h1,h2)).ToString());
__Check("True");

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
