// vybe-test: csharp/csharp_numerics_complex_arithmetic/complex_constants_zero_one_imaginaryone

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

__P(System.Numerics.Complex.Zero.Real.ToString(System.Globalization.CultureInfo.InvariantCulture));
__P(System.Numerics.Complex.One.Real.ToString(System.Globalization.CultureInfo.InvariantCulture));
__P(System.Numerics.Complex.ImaginaryOne.Imaginary.ToString(System.Globalization.CultureInfo.InvariantCulture));
__Check("0\n1\n1");
