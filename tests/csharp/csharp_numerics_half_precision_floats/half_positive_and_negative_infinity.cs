// vybe-test: csharp/csharp_numerics_half_precision_floats/half_positive_and_negative_infinity

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

__P(Half.IsPositiveInfinity(Half.PositiveInfinity).ToString());
__P(Half.IsNegativeInfinity(Half.NegativeInfinity).ToString());
__Check("True\nTrue");
