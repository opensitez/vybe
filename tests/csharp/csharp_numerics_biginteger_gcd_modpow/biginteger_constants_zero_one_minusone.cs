// vybe-test: csharp/csharp_numerics_biginteger_gcd_modpow/biginteger_constants_zero_one_minusone

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

__P(System.Numerics.BigInteger.Zero.ToString());
__P(System.Numerics.BigInteger.One.ToString());
__P(System.Numerics.BigInteger.MinusOne.ToString());
__Check("0\n1\n-1");
