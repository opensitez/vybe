// vybe-test: csharp/csharp_numerics_quaternion_plane/quaternion_dot_product

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

var q1 = new System.Numerics.Quaternion(1.0f, 0.0f, 0.0f, 0.0f);
var q2 = new System.Numerics.Quaternion(1.0f, 0.0f, 0.0f, 0.0f);
float dot = System.Numerics.Quaternion.Dot(q1, q2);
__P(dot.ToString(System.Globalization.CultureInfo.InvariantCulture));
__Check("1");
