// vybe-test: csharp/csharp_numerics_fused_multiply_add_math/math_tau_constant

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

double tau = Math.Tau;
__P((Math.Abs(tau - 2.0 * Math.PI) < 1e-9).ToString());
__Check("True");
