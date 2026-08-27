// vybe-test: csharp/csharp_numerics_half_precision_floats/half_icomparable_ordering

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

Half h1 = (Half)1.0f;
Half h2 = (Half)2.0f;
__P((h1.CompareTo(h2) < 0).ToString());
__Check("True");
