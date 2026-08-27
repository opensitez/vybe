// vybe-test: csharp/csharp_numerics_matrix3x2_matrix4x4/matrix4x4_create_look_at

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

var eye = new System.Numerics.Vector3(0, 0, 10);
var target = System.Numerics.Vector3.Zero;
var up = System.Numerics.Vector3.UnitY;
var view = System.Numerics.Matrix4x4.CreateLookAt(eye, target, up);
__P((view.M33 > 0).ToString());
__Check("True");
