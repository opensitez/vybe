// vybe-test: csharp/csharp_numerics_vector2_vector3_vector4/vector3_scalar_multiplication

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

var v = new System.Numerics.Vector3(2.0f, 3.0f, 4.0f);
var scaled = v * 2.0f;
__P(scaled.X.ToString(System.Globalization.CultureInfo.InvariantCulture));
__P(scaled.Z.ToString(System.Globalization.CultureInfo.InvariantCulture));
__Check("4\n8");
