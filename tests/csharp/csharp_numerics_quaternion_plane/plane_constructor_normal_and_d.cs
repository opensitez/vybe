// vybe-test: csharp/csharp_numerics_quaternion_plane/plane_constructor_normal_and_d

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

var normal = System.Numerics.Vector3.UnitY;
var plane = new System.Numerics.Plane(normal, -5.0f);
__P(plane.Normal.Y.ToString(System.Globalization.CultureInfo.InvariantCulture));
__P(plane.D.ToString(System.Globalization.CultureInfo.InvariantCulture));
__Check("1\n-5");
