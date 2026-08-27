// vybe-test: csharp/csharp_numerics_biginteger_gcd_modpow/biginteger_parse_large_string

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

var bi = System.Numerics.BigInteger.Parse("1234567890123456789012345678901234567890");
__P(bi.ToString());
__Check("1234567890123456789012345678901234567890");
