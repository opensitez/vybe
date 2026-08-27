// vybe-test: csharp/csharp_numerics_biginteger_gcd_modpow/biginteger_sign_property

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

__P(new System.Numerics.BigInteger(-50).Sign.ToString());
__P(new System.Numerics.BigInteger(0).Sign.ToString());
__P(new System.Numerics.BigInteger(50).Sign.ToString());
__Check("-1\n0\n1");
