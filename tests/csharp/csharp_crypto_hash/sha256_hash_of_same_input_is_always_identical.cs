// vybe-test: csharp/csharp_crypto_hash/sha256_hash_of_same_input_is_always_identical
// origin: languages/csharp/tests/csharp/test_csharp_crypto_hash.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

byte[] input=System.Text.Encoding.UTF8.GetBytes("hello");
var h1=System.Security.Cryptography.SHA256.HashData(input);
var h2=System.Security.Cryptography.SHA256.HashData(input);
__Check((System.MemoryExtensions.SequenceEqual<byte>(h1,h2)).ToString(), "True");
