// vybe-test: csharp/csharp_numerics_matrix3x2_matrix4x4/matrix4x4_multiplication_operator

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

var m1 = System.Numerics.Matrix4x4.CreateScale(2.0f);
var m2 = System.Numerics.Matrix4x4.CreateScale(3.0f);
var prod = m1 * m2;
__P(prod.M11.ToString(System.Globalization.CultureInfo.InvariantCulture));
__Check("6");
