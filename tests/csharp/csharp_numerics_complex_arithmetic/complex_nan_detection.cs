// vybe-test: csharp/csharp_numerics_complex_arithmetic/complex_nan_detection

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

var nan = new System.Numerics.Complex(double.NaN, 0);
__P(System.Numerics.Complex.IsNaN(nan).ToString());
__Check("True");
