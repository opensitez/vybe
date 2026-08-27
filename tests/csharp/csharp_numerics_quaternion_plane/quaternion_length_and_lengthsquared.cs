// vybe-test: csharp/csharp_numerics_quaternion_plane/quaternion_length_and_lengthsquared

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
__P(q.Length().ToString(System.Globalization.CultureInfo.InvariantCulture));
__P(q.LengthSquared().ToString(System.Globalization.CultureInfo.InvariantCulture));
__Check("1\n1");
