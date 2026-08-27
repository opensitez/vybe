// vybe-test: csharp/csharp_numerics_biginteger_gcd_modpow/biginteger_pow_method

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

var b = new System.Numerics.BigInteger(2);
var p = System.Numerics.BigInteger.Pow(b, 32);
__P(p.ToString());
__Check("4294967296");
