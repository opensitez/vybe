// vybe-test: csharp/csharp_numerics_fused_multiply_add_math/math_scaleb_double

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

double scaled = Math.ScaleB(1.0, 3); // 1.0 * 2^3 = 8
__P(scaled.ToString(System.Globalization.CultureInfo.InvariantCulture));
__Check("8");
