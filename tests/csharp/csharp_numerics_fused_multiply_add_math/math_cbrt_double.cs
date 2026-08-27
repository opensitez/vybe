// vybe-test: csharp/csharp_numerics_fused_multiply_add_math/math_cbrt_double

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

double cbrt = Math.Cbrt(27.0);
__P(cbrt.ToString(System.Globalization.CultureInfo.InvariantCulture));
__Check("3");
