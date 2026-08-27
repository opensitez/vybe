// vybe-test: csharp/csharp_numerics_vector2_vector3_vector4/vector2_in_generic_list

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

var list = new System.Collections.Generic.List<System.Numerics.Vector2>();
list.Add(System.Numerics.Vector2.One);
__P(list.Count.ToString());
__P(list[0].X.ToString(System.Globalization.CultureInfo.InvariantCulture));
__Check("1\n1");
