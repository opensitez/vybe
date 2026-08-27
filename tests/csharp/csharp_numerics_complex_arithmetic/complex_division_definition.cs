// vybe-test: csharp/csharp_numerics_complex_arithmetic/complex_division_definition

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

var c1 = new System.Numerics.Complex(-5.0, 10.0);
var c2 = new System.Numerics.Complex(1.0, 2.0);
var div = c1 / c2;
__P(div.Real.ToString(System.Globalization.CultureInfo.InvariantCulture));
__P(div.Imaginary.ToString(System.Globalization.CultureInfo.InvariantCulture));
__Check("3\n4");
