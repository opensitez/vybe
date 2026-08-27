// vybe-test: csharp/csharp_numerics_vector2_vector3_vector4/vector2_lerp_interpolation

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

var v1 = new System.Numerics.Vector2(0.0f, 0.0f);
var v2 = new System.Numerics.Vector2(10.0f, 20.0f);
var mid = System.Numerics.Vector2.Lerp(v1, v2, 0.5f);
__P(mid.X.ToString(System.Globalization.CultureInfo.InvariantCulture));
__P(mid.Y.ToString(System.Globalization.CultureInfo.InvariantCulture));
__Check("5\n10");
