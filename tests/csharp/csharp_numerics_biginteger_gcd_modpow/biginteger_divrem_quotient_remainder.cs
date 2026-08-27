// vybe-test: csharp/csharp_numerics_biginteger_gcd_modpow/biginteger_divrem_quotient_remainder

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

var a = new System.Numerics.BigInteger(100);
var b = new System.Numerics.BigInteger(30);
var q = System.Numerics.BigInteger.DivRem(a, b, out var r);
__P(q.ToString());
__P(r.ToString());
__Check("3\n10");
