// vybe-test: csharp/csharp_numerics_vector2_vector3_vector4/vector3_clamp_values

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

var v = new System.Numerics.Vector3(15.0f, -5.0f, 5.0f);
var min = new System.Numerics.Vector3(0.0f, 0.0f, 0.0f);
var max = new System.Numerics.Vector3(10.0f, 10.0f, 10.0f);
var clamped = System.Numerics.Vector3.Clamp(v, min, max);
__P(clamped.X.ToString(System.Globalization.CultureInfo.InvariantCulture));
__P(clamped.Y.ToString(System.Globalization.CultureInfo.InvariantCulture));
__Check("10\n0");
