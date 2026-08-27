// vybe-test: csharp/csharp_numerics_half_precision_floats/half_is_negative_check

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

Half neg = (Half)(-5.5f);
Half pos = (Half)5.5f;
__P(Half.IsNegative(neg).ToString());
__P(Half.IsNegative(pos).ToString());
__Check("True\nFalse");
