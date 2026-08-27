// vybe-test: csharp/csharp_numerics_fused_multiply_add_math/math_clamp_int32

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

int val = -5;
int clamped = Math.Clamp(val, 0, 100);
__P(clamped.ToString());
__Check("0");
