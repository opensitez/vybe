// vybe-test: csharp/csharp_crypto_hash/sha256_produces_32_byte_digest
// origin: languages/csharp/tests/csharp/test_csharp_crypto_hash.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

var hash=System.Security.Cryptography.SHA256.HashData(new byte[]{1,2,3});
__Check((hash.Length).ToString(), "32");
