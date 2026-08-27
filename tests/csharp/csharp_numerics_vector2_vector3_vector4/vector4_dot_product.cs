// vybe-test: csharp/csharp_numerics_vector2_vector3_vector4/vector4_dot_product

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

var v1 = new System.Numerics.Vector4(1.0f, 1.0f, 1.0f, 1.0f);
var v2 = new System.Numerics.Vector4(2.0f, 3.0f, 4.0f, 5.0f);
float dot = System.Numerics.Vector4.Dot(v1, v2);
__P(dot.ToString(System.Globalization.CultureInfo.InvariantCulture));
__Check("14");
