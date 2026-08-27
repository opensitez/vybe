// vybe-test: csharp/csharp_crypto_random_number_generator_bytes/rng_bytes_case_15

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

int val = System.Security.Cryptography.RandomNumberGenerator.GetInt32(1, 100);
__P((val >= 1 && val < 100).ToString());
__Check("True");
