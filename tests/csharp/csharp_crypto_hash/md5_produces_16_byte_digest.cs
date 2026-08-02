// vybe-test: csharp/csharp_crypto_hash/md5_produces_16_byte_digest
// origin: languages/csharp/tests/csharp/test_csharp_crypto_hash.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

using var md5=System.Security.Cryptography.MD5.Create();
byte[] hash=md5.ComputeHash(new byte[]{1,2,3});
__Check((hash.Length).ToString(), "16");
