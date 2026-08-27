// vybe-test: csharp/csharp_numerics_fused_multiply_add_math/math_fused_multiply_add_double

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

double res = Math.FusedMultiplyAdd(3.0, 4.0, 5.0); // 3*4 + 5 = 17
__P(res.ToString(System.Globalization.CultureInfo.InvariantCulture));
__Check("17");
