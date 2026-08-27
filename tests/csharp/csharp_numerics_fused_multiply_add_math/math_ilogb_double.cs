// vybe-test: csharp/csharp_numerics_fused_multiply_add_math/math_ilogb_double

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

int exp = Math.ILogB(8.0);
__P(exp.ToString());
__Check("3");
