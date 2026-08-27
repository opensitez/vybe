// vybe-test: csharp/csharp_numerics_complex_arithmetic/complex_exp_and_log

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
var exp = System.Numerics.Complex.Exp(z);
var log = System.Numerics.Complex.Log(System.Numerics.Complex.One);
__P(exp.Real.ToString(System.Globalization.CultureInfo.InvariantCulture));
__P(log.Real.ToString(System.Globalization.CultureInfo.InvariantCulture));
__Check("1\n0");
