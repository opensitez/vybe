// vybe-test: csharp/csharp_class_features/memberwise_clone_creates_shallow_copy
// origin: languages/csharp/tests/csharp/test_csharp_class_features.rs

using static __Harness;

var a=new Point{X=1,Y=2}
;
var b=(Point)a.Clone();
b.X=99;
__P((a.X).ToString());
__Check("1");

class Point:System.ICloneable{public int X,Y;public object Clone()=>MemberwiseClone();}

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
