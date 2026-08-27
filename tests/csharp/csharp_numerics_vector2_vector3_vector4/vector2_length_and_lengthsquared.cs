// vybe-test: csharp/csharp_numerics_vector2_vector3_vector4/vector2_length_and_lengthsquared

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

var v = new System.Numerics.Vector2(3.0f, 4.0f);
__P(v.Length().ToString(System.Globalization.CultureInfo.InvariantCulture));
__P(v.LengthSquared().ToString(System.Globalization.CultureInfo.InvariantCulture));
__Check("5\n25");
