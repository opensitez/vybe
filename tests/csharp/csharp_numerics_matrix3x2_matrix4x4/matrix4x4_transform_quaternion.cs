// vybe-test: csharp/csharp_numerics_matrix3x2_matrix4x4/matrix4x4_transform_quaternion

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

var q = System.Numerics.Quaternion.Identity;
var mat = System.Numerics.Matrix4x4.CreateFromQuaternion(q);
__P(mat.IsIdentity.ToString());
__Check("True");
