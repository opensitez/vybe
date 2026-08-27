// vybe-test: csharp/csharp_numerics_biginteger_gcd_modpow/biginteger_log_and_log10

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

var val = new System.Numerics.BigInteger(1000);
double log10 = System.Numerics.BigInteger.Log10(val);
__P(Math.Round(log10).ToString(System.Globalization.CultureInfo.InvariantCulture));
__Check("3");
