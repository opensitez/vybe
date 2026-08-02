// vybe-test: csharp/csharp_crypto_hash/sha1_produces_20_byte_digest
// origin: languages/csharp/tests/csharp/test_csharp_crypto_hash.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

var hash=System.Security.Cryptography.SHA1.HashData(new byte[]{0});
__Check((hash.Length).ToString(), "20");
