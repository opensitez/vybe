// vybe-test: csharp/csharp_numerics_matrix3x2_matrix4x4/matrix4x4_invert_matrix

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

var mat = System.Numerics.Matrix4x4.CreateScale(4.0f);
bool ok = System.Numerics.Matrix4x4.Invert(mat, out var inv);
__P(ok.ToString());
__P(inv.M11.ToString(System.Globalization.CultureInfo.InvariantCulture));
__Check("True\n0.25");
