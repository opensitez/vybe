// vybe-test: csharp/csharp_numerics_biginteger_gcd_modpow/biginteger_to_byte_array

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

var bi = new System.Numerics.BigInteger(255);
byte[] bytes = bi.ToByteArray();
__P(bytes.Length.ToString());
__P(bytes[0].ToString());
__Check("2\n255");
