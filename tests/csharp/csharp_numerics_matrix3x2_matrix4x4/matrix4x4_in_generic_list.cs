// vybe-test: csharp/csharp_numerics_matrix3x2_matrix4x4/matrix4x4_in_generic_list

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

var list = new System.Collections.Generic.List<System.Numerics.Matrix4x4>();
list.Add(System.Numerics.Matrix4x4.Identity);
__P(list.Count.ToString());
__P(list[0].IsIdentity.ToString());
__Check("1\nTrue");
