// vybe-test: csharp/csharp_numerics_vector2_vector3_vector4/vector2_equality_and_hashcode

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

var v1 = new System.Numerics.Vector2(1.5f, 2.5f);
var v2 = new System.Numerics.Vector2(1.5f, 2.5f);
__P((v1 == v2).ToString());
__P((v1.GetHashCode() == v2.GetHashCode()).ToString());
__Check("True\nTrue");
