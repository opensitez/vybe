// vybe-test: csharp/csharp_numerics_half_precision_floats/half_zero_and_negative_zero

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

Half z1 = (Half)0.0f;
Half z2 = (Half)(-0.0f);
__P((z1 == z2).ToString());
__Check("True");
