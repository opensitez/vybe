// vybe-test: csharp/csharp_numerics_fused_multiply_add_math/math_sin_cos_pi_half

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

(double sin, double cos) = Math.SinCos(Math.PI / 2.0);
__P(Math.Round(sin, 2).ToString(System.Globalization.CultureInfo.InvariantCulture));
__P(Math.Round(cos, 2).ToString(System.Globalization.CultureInfo.InvariantCulture));
__Check("1\n0");
