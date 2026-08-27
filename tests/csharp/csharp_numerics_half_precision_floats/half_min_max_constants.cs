// vybe-test: csharp/csharp_numerics_half_precision_floats/half_min_max_constants

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

__P(((double)Half.MinValue < -65500).ToString());
__P(((double)Half.MaxValue > 65500).ToString());
__Check("True\nTrue");
