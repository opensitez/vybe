// vybe-test: csharp/csharp_numerics_biginteger_gcd_modpow/biginteger_equality_and_hashcode

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

var a = new System.Numerics.BigInteger(12345);
var b = new System.Numerics.BigInteger(12345);
__P((a == b).ToString());
__P((a.GetHashCode() == b.GetHashCode()).ToString());
__Check("True\nTrue");
