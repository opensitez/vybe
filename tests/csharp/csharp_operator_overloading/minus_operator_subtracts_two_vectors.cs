// vybe-test: csharp/csharp_operator_overloading/minus_operator_subtracts_two_vectors
// origin: languages/csharp/tests/csharp/test_csharp_operator_overloading.rs

using static __Harness;

var v=new Vec{X=5,Y=3}
-new Vec{X=2,Y=1}
;
__P((v.X).ToString());
__Check("3");

struct Vec{public int X,Y;
public static Vec operator-(Vec a,Vec b)=>new Vec{X=a.X-b.X,Y=a.Y-b.Y};}

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
