// vybe-test: csharp/csharp_numerics_vector2_vector3_vector4/vector3_reflect_vector

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

var v = new System.Numerics.Vector3(0.0f, -1.0f, 0.0f);
var normal = new System.Numerics.Vector3(0.0f, 1.0f, 0.0f);
var refVector = System.Numerics.Vector3.Reflect(v, normal);
__P(refVector.Y.ToString(System.Globalization.CultureInfo.InvariantCulture));
__Check("1");
