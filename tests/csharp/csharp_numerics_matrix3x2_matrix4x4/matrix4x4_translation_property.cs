// vybe-test: csharp/csharp_numerics_matrix3x2_matrix4x4/matrix4x4_translation_property

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

var mat = System.Numerics.Matrix4x4.CreateTranslation(1.0f, 2.0f, 3.0f);
var trans = mat.Translation;
__P(trans.X.ToString(System.Globalization.CultureInfo.InvariantCulture));
__P(trans.Y.ToString(System.Globalization.CultureInfo.InvariantCulture));
__P(trans.Z.ToString(System.Globalization.CultureInfo.InvariantCulture));
__Check("1\n2\n3");
