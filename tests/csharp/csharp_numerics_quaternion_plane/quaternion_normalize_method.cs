// vybe-test: csharp/csharp_numerics_quaternion_plane/quaternion_normalize_method

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

var q = new System.Numerics.Quaternion(0.0f, 0.0f, 0.0f, 5.0f);
var norm = System.Numerics.Quaternion.Normalize(q);
__P(norm.W.ToString(System.Globalization.CultureInfo.InvariantCulture));
__Check("1");
