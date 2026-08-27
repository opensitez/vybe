// vybe-test: csharp/csharp_numerics_biginteger_gcd_modpow/biginteger_shifts_left_and_right

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

var a = new System.Numerics.BigInteger(1);
var shifted = a << 10;
var back = shifted >> 10;
__P(shifted.ToString());
__P(back.ToString());
__Check("1024\n1");
