// vybe-test: csharp/csharp_numerics_biginteger_gcd_modpow/biginteger_compare_to_ordering

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

var a = new System.Numerics.BigInteger(10);
var b = new System.Numerics.BigInteger(20);
__P((a.CompareTo(b) < 0).ToString());
__Check("True");
