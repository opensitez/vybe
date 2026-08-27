// vybe-test: csharp/csharp_numerics_biginteger_gcd_modpow/biginteger_greatest_common_divisor

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

var a = new System.Numerics.BigInteger(54);
var b = new System.Numerics.BigInteger(24);
var gcd = System.Numerics.BigInteger.GreatestCommonDivisor(a, b);
__P(gcd.ToString());
__Check("6");
