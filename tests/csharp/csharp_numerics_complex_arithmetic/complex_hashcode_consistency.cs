// vybe-test: csharp/csharp_numerics_complex_arithmetic/complex_hashcode_consistency

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

var c1 = new System.Numerics.Complex(7.0, 8.0);
var c2 = new System.Numerics.Complex(7.0, 8.0);
__P((c1.GetHashCode() == c2.GetHashCode()).ToString());
__Check("True");
