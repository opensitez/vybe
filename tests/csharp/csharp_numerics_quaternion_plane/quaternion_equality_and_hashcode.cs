// vybe-test: csharp/csharp_numerics_quaternion_plane/quaternion_equality_and_hashcode

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
__P((q1 == q2).ToString());
__P((q1.GetHashCode() == q2.GetHashCode()).ToString());
__Check("True\nTrue");
