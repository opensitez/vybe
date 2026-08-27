// vybe-test: csharp/csharp_numerics_complex_arithmetic/complex_negation_operator

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

var c = new System.Numerics.Complex(3.0, -4.0);
var neg = -c;
__P(neg.Real.ToString(System.Globalization.CultureInfo.InvariantCulture));
__P(neg.Imaginary.ToString(System.Globalization.CultureInfo.InvariantCulture));
__Check("-3\n4");
