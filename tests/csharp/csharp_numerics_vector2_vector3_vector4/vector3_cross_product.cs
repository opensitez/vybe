// vybe-test: csharp/csharp_numerics_vector2_vector3_vector4/vector3_cross_product

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

var v1 = System.Numerics.Vector3.UnitX;
var v2 = System.Numerics.Vector3.UnitY;
var cross = System.Numerics.Vector3.Cross(v1, v2);
__P((cross == System.Numerics.Vector3.UnitZ).ToString());
__Check("True");
