// vybe-test: csharp/csharp_numerics_biginteger_gcd_modpow/biginteger_modpow_calculation

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

var val = new System.Numerics.BigInteger(4);
var exp = new System.Numerics.BigInteger(13);
var mod = new System.Numerics.BigInteger(497);
var res = System.Numerics.BigInteger.ModPow(val, exp, mod);
__P(res.ToString());
__Check("445");
