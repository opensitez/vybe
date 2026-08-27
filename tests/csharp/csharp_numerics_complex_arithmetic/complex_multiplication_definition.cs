// vybe-test: csharp/csharp_numerics_complex_arithmetic/complex_multiplication_definition

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

var c1 = new System.Numerics.Complex(1.0, 2.0);
var c2 = new System.Numerics.Complex(3.0, 4.0);
var prod = c1 * c2; // (1*3 - 2*4) + (1*4 + 2*3)i = -5 + 10i
__P(prod.Real.ToString(System.Globalization.CultureInfo.InvariantCulture));
__P(prod.Imaginary.ToString(System.Globalization.CultureInfo.InvariantCulture));
__Check("-5\n10");
