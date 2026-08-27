// vybe-test: csharp/csharp_crypto_sha256_sha512_hashes/sha256_hash_case_20

string __buf = "";

void __P(string s) {
    __buf = __buf + s + "\n";
}

void __Pr(string s) {
    __buf = __buf + s;
}

void __Check(string want) {
    if (__buf != want && __buf != want + "\n") {
        Console.WriteLine("FAIL: want [" + want + "] got [" + __buf + "]");
        throw new Exception("assertion failed");
    }
}

byte[] data = System.Text.Encoding.UTF8.GetBytes("Data_20");
byte[] hash = System.Security.Cryptography.SHA256.HashData(data);
__P(hash.Length.ToString());
__Check("32");
