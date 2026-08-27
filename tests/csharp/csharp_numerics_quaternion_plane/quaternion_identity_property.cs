// vybe-test: csharp/csharp_numerics_quaternion_plane/quaternion_identity_property

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
__P(q.IsIdentity.ToString());
__P(q.W.ToString(System.Globalization.CultureInfo.InvariantCulture));
__P(q.X.ToString(System.Globalization.CultureInfo.InvariantCulture));
__Check("True\n1\n0");
