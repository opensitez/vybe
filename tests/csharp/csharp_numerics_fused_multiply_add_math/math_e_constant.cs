// vybe-test: csharp/csharp_numerics_fused_multiply_add_math/math_e_constant

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

double e = Math.E;
__P((e > 2.71 && e < 2.72).ToString());
__Check("True");
