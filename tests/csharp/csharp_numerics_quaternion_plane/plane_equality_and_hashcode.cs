// vybe-test: csharp/csharp_numerics_quaternion_plane/plane_equality_and_hashcode

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

var p1 = new System.Numerics.Plane(System.Numerics.Vector3.UnitX, 1.0f);
var p2 = new System.Numerics.Plane(System.Numerics.Vector3.UnitX, 1.0f);
__P((p1 == p2).ToString());
__P((p1.GetHashCode() == p2.GetHashCode()).ToString());
__Check("True\nTrue");
