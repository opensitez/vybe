// vybe-test: csharp/csharp_numerics_complex_arithmetic/complex_pow_imaginary_unit

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

var i = System.Numerics.Complex.ImaginaryOne;
var i2 = System.Numerics.Complex.Pow(i, 2.0);
__P(Math.Round(i2.Real, 2).ToString(System.Globalization.CultureInfo.InvariantCulture));
__P(Math.Round(i2.Imaginary, 2).ToString(System.Globalization.CultureInfo.InvariantCulture));
__Check("-1\n0");
