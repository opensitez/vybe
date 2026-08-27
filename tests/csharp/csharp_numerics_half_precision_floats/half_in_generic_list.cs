// vybe-test: csharp/csharp_numerics_half_precision_floats/half_in_generic_list

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

var list = new System.Collections.Generic.List<Half>();
list.Add((Half)1.0f);
list.Add((Half)2.0f);
__P(list.Count.ToString());
__P(list[1].ToString(System.Globalization.CultureInfo.InvariantCulture));
__Check("2\n2");
