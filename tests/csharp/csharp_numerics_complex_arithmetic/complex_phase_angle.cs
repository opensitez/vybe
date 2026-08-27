// vybe-test: csharp/csharp_numerics_complex_arithmetic/complex_phase_angle

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

var c = new System.Numerics.Complex(0.0, 1.0);
double halfPi = Math.PI / 2.0;
__P((Math.Abs(c.Phase - halfPi) < 1e-9).ToString());
__Check("True");
