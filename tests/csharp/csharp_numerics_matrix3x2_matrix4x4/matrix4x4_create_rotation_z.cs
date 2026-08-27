// vybe-test: csharp/csharp_numerics_matrix3x2_matrix4x4/matrix4x4_create_rotation_z

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

var mat = System.Numerics.Matrix4x4.CreateRotationZ(0.0f);
__P(mat.IsIdentity.ToString());
__Check("True");
