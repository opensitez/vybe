// vybe-test: csharp/csharp_crypto_hash/sha256_hash_of_same_input_is_always_identical
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

byte[] input=System.Text.Encoding.UTF8.GetBytes("hello");
var h1=System.Security.Cryptography.SHA256.HashData(input);
var h2=System.Security.Cryptography.SHA256.HashData(input);
__P((System.MemoryExtensions.SequenceEqual<byte>(h1,h2)).ToString());
__Check("True");
