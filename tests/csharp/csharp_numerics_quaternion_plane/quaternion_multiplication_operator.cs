// vybe-test: csharp/csharp_numerics_quaternion_plane/quaternion_multiplication_operator

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

var q1 = System.Numerics.Quaternion.Identity;
var q2 = System.Numerics.Quaternion.Identity;
var prod = q1 * q2;
__P(prod.IsIdentity.ToString());
__Check("True");
