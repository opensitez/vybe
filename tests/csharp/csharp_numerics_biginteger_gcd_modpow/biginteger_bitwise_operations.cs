// vybe-test: csharp/csharp_numerics_biginteger_gcd_modpow/biginteger_bitwise_operations

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

var a = new System.Numerics.BigInteger(0xFF);
var b = new System.Numerics.BigInteger(0x0F);
__P((a & b).ToString());
__P((a | b).ToString());
__P((a ^ b).ToString());
__Check("15\n255\n240");
