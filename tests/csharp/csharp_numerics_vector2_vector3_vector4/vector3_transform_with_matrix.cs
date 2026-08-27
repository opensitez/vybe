// vybe-test: csharp/csharp_numerics_vector2_vector3_vector4/vector3_transform_with_matrix

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

var v = System.Numerics.Vector3.UnitX;
var mat = System.Numerics.Matrix4x4.CreateTranslation(5.0f, 0.0f, 0.0f);
var trans = System.Numerics.Vector3.Transform(v, mat);
__P(trans.X.ToString(System.Globalization.CultureInfo.InvariantCulture));
__Check("6");
