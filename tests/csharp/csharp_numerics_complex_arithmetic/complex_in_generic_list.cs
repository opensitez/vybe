// vybe-test: csharp/csharp_numerics_complex_arithmetic/complex_in_generic_list

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

var list = new System.Collections.Generic.List<System.Numerics.Complex>();
list.Add(new System.Numerics.Complex(10.0, 20.0));
__P(list.Count.ToString());
__P(list[0].Real.ToString(System.Globalization.CultureInfo.InvariantCulture));
__Check("1\n10");
