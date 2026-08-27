// vybe-test: csharp/csharp_numerics_half_precision_floats/half_explicit_cast_from_double

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

double d = -10.5;
Half h = (Half)d;
__P(h.ToString(System.Globalization.CultureInfo.InvariantCulture));
__Check("-10.5");
