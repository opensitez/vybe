// vybe-test: csharp/csharp_numerics_half_precision_floats/half_parse_and_tostring_invariants

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

Half h = Half.Parse("1.5", System.Globalization.CultureInfo.InvariantCulture);
__P(h.ToString(System.Globalization.CultureInfo.InvariantCulture));
__Check("1.5");
