// vybe-test: csharp/csharp_numerics_matrix3x2_matrix4x4/matrix3x2_create_scale

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

var mat = System.Numerics.Matrix3x2.CreateScale(2.0f, 3.0f);
__P(mat.M11.ToString(System.Globalization.CultureInfo.InvariantCulture));
__P(mat.M22.ToString(System.Globalization.CultureInfo.InvariantCulture));
__Check("2\n3");
