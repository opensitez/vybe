// vybe-test: csharp/csharp_numerics_complex_arithmetic/complex_sin_and_cos

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

var z = System.Numerics.Complex.Zero;
var sin = System.Numerics.Complex.Sin(z);
var cos = System.Numerics.Complex.Cos(z);
__P(sin.Real.ToString(System.Globalization.CultureInfo.InvariantCulture));
__P(cos.Real.ToString(System.Globalization.CultureInfo.InvariantCulture));
__Check("0\n1");
