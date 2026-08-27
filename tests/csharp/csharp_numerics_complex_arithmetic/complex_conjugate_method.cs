// vybe-test: csharp/csharp_numerics_complex_arithmetic/complex_conjugate_method

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

var c = new System.Numerics.Complex(5.0, -2.0);
var conj = System.Numerics.Complex.Conjugate(c);
__P(conj.Real.ToString(System.Globalization.CultureInfo.InvariantCulture));
__P(conj.Imaginary.ToString(System.Globalization.CultureInfo.InvariantCulture));
__Check("5\n2");
