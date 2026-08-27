// vybe-test: csharp/csharp_numerics_biginteger_gcd_modpow/biginteger_from_byte_array

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

byte[] bytes = new byte[] { 0x00, 0x01 }; // 256 in little endian
var bi = new System.Numerics.BigInteger(bytes);
__P(bi.ToString());
__Check("256");
