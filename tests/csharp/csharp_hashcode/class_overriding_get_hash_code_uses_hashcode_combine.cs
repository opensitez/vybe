// vybe-test: csharp/csharp_hashcode/class_overriding_get_hash_code_uses_hashcode_combine
// origin: languages/csharp/tests/csharp/test_csharp_hashcode.rs

string __buf = "";

void __P(string s) {
    __buf = __buf + s + "\n";
}

void __Pr(string s) {
    __buf = __buf + s;
}

// The final WriteLine contributes a trailing newline that the expected line
// vector never carried, so BOTH forms are accepted.
void __Check(string want) {
    if (__buf != want && __buf != want + "\n") {
        Console.WriteLine("FAIL: want [" + want + "] got [" + __buf + "]");
        throw new Exception("assertion failed");
    }
}

class Point{
    public int X,Y;
    public override int GetHashCode()=>System.HashCode.Combine(X,Y);
}
var p1=new Point{X=1,Y=2};
var p2=new Point{X=1,Y=2};
__P((p1.GetHashCode()==p2.GetHashCode()).ToString());
__Check("True");
