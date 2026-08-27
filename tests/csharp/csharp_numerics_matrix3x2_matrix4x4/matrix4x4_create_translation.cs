// vybe-test: csharp/csharp_numerics_matrix3x2_matrix4x4/matrix4x4_create_translation

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

var mat = System.Numerics.Matrix4x4.CreateTranslation(5.0f, 10.0f, 15.0f);
__P(mat.M41.ToString(System.Globalization.CultureInfo.InvariantCulture));
__P(mat.M42.ToString(System.Globalization.CultureInfo.InvariantCulture));
__P(mat.M43.ToString(System.Globalization.CultureInfo.InvariantCulture));
__Check("5\n10\n15");
