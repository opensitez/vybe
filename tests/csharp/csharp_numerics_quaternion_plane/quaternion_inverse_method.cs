// vybe-test: csharp/csharp_numerics_quaternion_plane/quaternion_inverse_method

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
var inv = System.Numerics.Quaternion.Inverse(q);
__P(inv.IsIdentity.ToString());
__Check("True");
