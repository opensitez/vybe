// vybe-test: csharp/csharp_numerics_quaternion_plane/plane_transform_matrix

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

var plane = new System.Numerics.Plane(System.Numerics.Vector3.UnitZ, 0.0f);
var mat = System.Numerics.Matrix4x4.CreateTranslation(0.0f, 0.0f, 5.0f);
var trans = System.Numerics.Plane.Transform(plane, mat);
__P(trans.Normal.Z.ToString(System.Globalization.CultureInfo.InvariantCulture));
__Check("1");
