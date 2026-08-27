// vybe-test: csharp/csharp_numerics_quaternion_plane/plane_create_from_vertices

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

var p1 = new System.Numerics.Vector3(0, 0, 0);
var p2 = new System.Numerics.Vector3(1, 0, 0);
var p3 = new System.Numerics.Vector3(0, 1, 0);
var plane = System.Numerics.Plane.CreateFromVertices(p1, p2, p3);
__P(plane.Normal.Z.ToString(System.Globalization.CultureInfo.InvariantCulture));
__Check("1");
