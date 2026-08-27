// vybe-test: csharp/csharp_numerics_complex_arithmetic/complex_polar_coordinates_factory

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

var polar = System.Numerics.Complex.FromPolarCoordinates(5.0, 0.0);
__P(polar.Real.ToString(System.Globalization.CultureInfo.InvariantCulture));
__P(polar.Imaginary.ToString(System.Globalization.CultureInfo.InvariantCulture));
__Check("5\n0");
