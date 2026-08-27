// vybe-test: csharp/csharp_numerics_matrix3x2_matrix4x4/matrix3x2_create_translation

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

var mat = System.Numerics.Matrix3x2.CreateTranslation(10.0f, 20.0f);
__P(mat.M31.ToString(System.Globalization.CultureInfo.InvariantCulture));
__P(mat.M32.ToString(System.Globalization.CultureInfo.InvariantCulture));
__Check("10\n20");
