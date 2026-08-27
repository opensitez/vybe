// vybe-test: csharp/csharp_numerics_vector2_vector3_vector4/vector2_normalize_method

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

var v = new System.Numerics.Vector2(0.0f, 5.0f);
var norm = System.Numerics.Vector2.Normalize(v);
__P(norm.X.ToString(System.Globalization.CultureInfo.InvariantCulture));
__P(norm.Y.ToString(System.Globalization.CultureInfo.InvariantCulture));
__Check("0\n1");
