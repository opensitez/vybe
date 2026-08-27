// vybe-test: csharp/csharp_numerics_complex_arithmetic/complex_equality_and_inequality

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

var c1 = new System.Numerics.Complex(2.5, 3.5);
var c2 = new System.Numerics.Complex(2.5, 3.5);
var c3 = new System.Numerics.Complex(2.5, 4.0);
__P((c1 == c2).ToString());
__P((c1 != c3).ToString());
__Check("True\nTrue");
