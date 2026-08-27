// vybe-test: csharp/csharp_hashcode/class_overriding_get_hash_code_uses_hashcode_combine
// origin: languages/csharp/tests/csharp/test_csharp_hashcode.rs

using static __Harness;

var p1=new Point{X=1,Y=2}
;
var p2=new Point{X=1,Y=2}
;
__P((p1.GetHashCode()==p2.GetHashCode()).ToString());
__Check("True");

class Point{
    public int X,Y;
    public override int GetHashCode()=>System.HashCode.Combine(X,Y);
}

public static class __Harness {
    public static string __buf = "";
    public static void __P(string s) { __buf = __buf + s + "\n"; }
    public static void __Pr(string s) { __buf = __buf + s; }
    public static void __Check(string want) {
        if (__buf != want && __buf != want + "\n") {
            Console.WriteLine("FAIL: want [" + want + "] got [" + __buf + "]");
            throw new Exception("assertion failed");
        }
    }
}
