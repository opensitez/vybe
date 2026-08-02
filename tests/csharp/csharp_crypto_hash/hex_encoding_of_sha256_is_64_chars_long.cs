// vybe-test: csharp/csharp_crypto_hash/hex_encoding_of_sha256_is_64_chars_long
// origin: languages/csharp/tests/csharp/test_csharp_crypto_hash.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

var hash=System.Security.Cryptography.SHA256.HashData(System.Text.Encoding.UTF8.GetBytes("test"));
string hex=System.Convert.ToHexString(hash);
__Check((hex.Length).ToString(), "64");
