// vybe-test: csharp/csharp_numerics_half_precision_floats/half_equality_comparison

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

Half h1 = (Half)3.0f;
Half h2 = (Half)3.0f;
Half h3 = (Half)4.0f;
__P((h1 == h2).ToString());
__P((h1 != h3).ToString());
__Check("True\nTrue");
