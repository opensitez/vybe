// vybe-test: csharp/csharp_numerics_quaternion_plane/plane_normalize_method

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

var plane = new System.Numerics.Plane(new System.Numerics.Vector3(0.0f, 5.0f, 0.0f), 10.0f);
var norm = System.Numerics.Plane.Normalize(plane);
__P(norm.Normal.Y.ToString(System.Globalization.CultureInfo.InvariantCulture));
__P(norm.D.ToString(System.Globalization.CultureInfo.InvariantCulture));
__Check("1\n2");
