// vybe-test: csharp/csharp_numerics_complex_arithmetic/complex_is_finite_and_infinity

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

var c = new System.Numerics.Complex(1.0, 2.0);
var inf = new System.Numerics.Complex(double.PositiveInfinity, 0);
__P(System.Numerics.Complex.IsFinite(c).ToString());
__P(System.Numerics.Complex.IsInfinity(inf).ToString());
__Check("True\nTrue");
