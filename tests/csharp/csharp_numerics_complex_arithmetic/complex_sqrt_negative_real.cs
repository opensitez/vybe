// vybe-test: csharp/csharp_numerics_complex_arithmetic/complex_sqrt_negative_real

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

var c = new System.Numerics.Complex(-4.0, 0.0);
var sqrt = System.Numerics.Complex.Sqrt(c);
__P(sqrt.Real.ToString(System.Globalization.CultureInfo.InvariantCulture));
__P(sqrt.Imaginary.ToString(System.Globalization.CultureInfo.InvariantCulture));
__Check("0\n2");
