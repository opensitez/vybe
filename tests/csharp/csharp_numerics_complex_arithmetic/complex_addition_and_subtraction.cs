// vybe-test: csharp/csharp_numerics_complex_arithmetic/complex_addition_and_subtraction

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
var sum = c1 + c2;
var diff = c2 - c1;
__P(sum.Real.ToString(System.Globalization.CultureInfo.InvariantCulture));
__P(sum.Imaginary.ToString(System.Globalization.CultureInfo.InvariantCulture));
__P(diff.Real.ToString(System.Globalization.CultureInfo.InvariantCulture));
__P(diff.Imaginary.ToString(System.Globalization.CultureInfo.InvariantCulture));
__Check("4\n6\n2\n2");
