// vybe-test: csharp/csharp_numerics_fused_multiply_add_math/math_fused_multiply_add_float

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

float res = MathF.FusedMultiplyAdd(2.0f, 3.0f, 4.0f); // 2*3 + 4 = 10
__P(res.ToString(System.Globalization.CultureInfo.InvariantCulture));
__Check("10");
