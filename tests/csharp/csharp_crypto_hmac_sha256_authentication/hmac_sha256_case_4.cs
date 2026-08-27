// vybe-test: csharp/csharp_crypto_hmac_sha256_authentication/hmac_sha256_case_4

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

byte[] key = new byte[32];
byte[] data = System.Text.Encoding.UTF8.GetBytes("Msg_4");
byte[] hmac = System.Security.Cryptography.HMACSHA256.HashData(key, data);
__P(hmac.Length.ToString());
__Check("32");
