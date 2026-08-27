// vybe-test: csharp/csharp_numerics_quaternion_plane/plane_dot_normal

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

var plane = new System.Numerics.Plane(System.Numerics.Vector3.UnitY, 5.0f);
var v = new System.Numerics.Vector3(0.0f, 2.0f, 0.0f);
float dot = System.Numerics.Plane.DotNormal(plane, v);
__P(dot.ToString(System.Globalization.CultureInfo.InvariantCulture));
__Check("2");
