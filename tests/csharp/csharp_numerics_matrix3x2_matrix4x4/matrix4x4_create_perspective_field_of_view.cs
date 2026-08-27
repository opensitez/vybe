// vybe-test: csharp/csharp_numerics_matrix3x2_matrix4x4/matrix4x4_create_perspective_field_of_view

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

var fov = System.Numerics.Matrix4x4.CreatePerspectiveFieldOfView((float)Math.PI / 4.0f, 1.0f, 1.0f, 100.0f);
__P((fov.M11 > 0).ToString());
__Check("True");
