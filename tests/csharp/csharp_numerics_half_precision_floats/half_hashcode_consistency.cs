// vybe-test: csharp/csharp_numerics_half_precision_floats/half_hashcode_consistency

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

Half h1 = (Half)99.5f;
Half h2 = (Half)99.5f;
__P((h1.GetHashCode() == h2.GetHashCode()).ToString());
__Check("True");
