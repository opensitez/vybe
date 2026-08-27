// vybe-test: csharp/csharp_numerics_biginteger_gcd_modpow/biginteger_is_even_and_is_power_of_two

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

var even = new System.Numerics.BigInteger(64);
var odd = new System.Numerics.BigInteger(65);
__P(even.IsEven.ToString());
__P(even.IsPowerOfTwo.ToString());
__P(odd.IsEven.ToString());
__P(odd.IsPowerOfTwo.ToString());
__Check("True\nTrue\nFalse\nFalse");
