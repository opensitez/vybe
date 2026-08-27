// vybe-test: csharp/csharp_numerics_int128_uint128/int128_explicit_cast_to_double

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

Int128 v = (Int128)1000;
double d = (double)v;
__P(d.ToString(System.Globalization.CultureInfo.InvariantCulture));
__Check("1000");
