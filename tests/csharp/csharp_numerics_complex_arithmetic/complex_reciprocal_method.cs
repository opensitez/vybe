// vybe-test: csharp/csharp_numerics_complex_arithmetic/complex_reciprocal_method

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

var c = new System.Numerics.Complex(2.0, 0.0);
var rec = System.Numerics.Complex.Reciprocal(c);
__P(rec.Real.ToString(System.Globalization.CultureInfo.InvariantCulture));
__Check("0.5");
