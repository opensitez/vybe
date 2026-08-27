// vybe-test: csharp/csharp_numerics_biginteger_gcd_modpow/biginteger_abs_and_negate

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

var neg = new System.Numerics.BigInteger(-999);
var abs = System.Numerics.BigInteger.Abs(neg);
var back = System.Numerics.BigInteger.Negate(abs);
__P(abs.ToString());
__P(back.ToString());
__Check("999\n-999");
