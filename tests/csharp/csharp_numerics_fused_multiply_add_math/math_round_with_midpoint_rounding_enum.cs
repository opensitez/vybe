// vybe-test: csharp/csharp_numerics_fused_multiply_add_math/math_round_with_midpoint_rounding_enum

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

double r1 = Math.Round(2.5, MidpointRounding.ToEven);
double r2 = Math.Round(2.5, MidpointRounding.AwayFromZero);
__P(r1.ToString(System.Globalization.CultureInfo.InvariantCulture));
__P(r2.ToString(System.Globalization.CultureInfo.InvariantCulture));
__Check("2\n3");
