// vybe-test: csharp/csharp_numerics_half_precision_floats/half_explicit_cast_to_float

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

Half h = (Half)7.5f;
float f = (float)h;
__P(f.ToString(System.Globalization.CultureInfo.InvariantCulture));
__Check("7.5");
