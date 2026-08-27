// vybe-test: csharp/csharp_numerics_fused_multiply_add_math/math_bigmul_int32

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

long big = Math.BigMul(1000000, 1000000);
__P(big.ToString());
__Check("1000000000000");
