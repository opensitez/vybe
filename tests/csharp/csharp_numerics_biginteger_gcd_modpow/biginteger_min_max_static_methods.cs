// vybe-test: csharp/csharp_numerics_biginteger_gcd_modpow/biginteger_min_max_static_methods

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
var b = new System.Numerics.BigInteger(200);
__P(System.Numerics.BigInteger.Min(a, b).ToString());
__P(System.Numerics.BigInteger.Max(a, b).ToString());
__Check("100\n200");
