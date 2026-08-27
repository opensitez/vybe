// vybe-test: csharp/csharp_numerics_quaternion_plane/quaternion_conjugate_method

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

var q = new System.Numerics.Quaternion(1.0f, 2.0f, 3.0f, 4.0f);
var conj = System.Numerics.Quaternion.Conjugate(q);
__P(conj.X.ToString(System.Globalization.CultureInfo.InvariantCulture));
__P(conj.W.ToString(System.Globalization.CultureInfo.InvariantCulture));
__Check("-1\n4");
