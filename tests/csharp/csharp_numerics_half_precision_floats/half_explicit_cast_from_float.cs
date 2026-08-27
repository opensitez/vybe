// vybe-test: csharp/csharp_numerics_half_precision_floats/half_explicit_cast_from_float

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

float f = 42.25f;
Half h = (Half)f;
__P(h.ToString(System.Globalization.CultureInfo.InvariantCulture));
__Check("42.25");
