// vybe-test: csharp/csharp_numerics_fused_multiply_add_math/math_reciprocal_sqrt_double

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

double val = 16.0;
double rsqrt = 1.0 / Math.Sqrt(val);
__P(rsqrt.ToString(System.Globalization.CultureInfo.InvariantCulture));
__Check("0.25");
