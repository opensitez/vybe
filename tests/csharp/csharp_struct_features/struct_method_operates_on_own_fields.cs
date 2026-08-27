// vybe-test: csharp/csharp_struct_features/struct_method_operates_on_own_fields
// origin: languages/csharp/tests/csharp/test_csharp_struct_features.rs

using static __Harness;

var v = new Vector { X=3, Y=4 }
;
__P((v.Length()).ToString());
__Check("5");

struct Vector { public double X,Y; public double Length() => System.Math.Sqrt(X*X+Y*Y); }

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
