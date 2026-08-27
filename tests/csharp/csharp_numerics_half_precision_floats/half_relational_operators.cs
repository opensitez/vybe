// vybe-test: csharp/csharp_numerics_half_precision_floats/half_relational_operators

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

Half h1 = (Half)2.5f;
Half h2 = (Half)5.0f;
__P((h1 < h2).ToString());
__P((h2 > h1).ToString());
__Check("True\nTrue");
