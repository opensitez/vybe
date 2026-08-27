// vybe-test: csharp/csharp_numerics_fused_multiply_add_math/math_divrem_int64

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

long rem;
long quot = Math.DivRem(25L, 4L, out rem);
__P(quot.ToString());
__P(rem.ToString());
__Check("6\n1");
