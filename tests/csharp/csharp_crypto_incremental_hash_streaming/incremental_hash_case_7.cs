// vybe-test: csharp/csharp_crypto_incremental_hash_streaming/incremental_hash_case_7

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

using var inc = System.Security.Cryptography.IncrementalHash.CreateHash(System.Security.Cryptography.HashAlgorithmName.SHA256);
inc.AppendData(new byte[] { (byte)7 });
byte[] hash = inc.GetHashAndReset();
__P(hash.Length.ToString());
__Check("32");
