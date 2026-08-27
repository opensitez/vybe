// vybe-test: csharp/csharp_numerics_biginteger_gcd_modpow/biginteger_tryparse_hex_style

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

bool ok = System.Numerics.BigInteger.TryParse("1A", System.Globalization.NumberStyles.HexNumber, null, out var res);
__P(ok.ToString());
__P(res.ToString());
__Check("True\n26");
