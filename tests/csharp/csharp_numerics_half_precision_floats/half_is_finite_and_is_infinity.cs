// vybe-test: csharp/csharp_numerics_half_precision_floats/half_is_finite_and_is_infinity

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

Half finite = (Half)100.0f;
Half inf = Half.PositiveInfinity;
__P(Half.IsFinite(finite).ToString());
__P(Half.IsInfinity(inf).ToString());
__Check("True\nTrue");
