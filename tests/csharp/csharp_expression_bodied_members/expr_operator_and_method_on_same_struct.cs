// vybe-test: csharp/csharp_expression_bodied_members/expr_operator_and_method_on_same_struct
// origin: languages/csharp/tests/csharp/test_csharp_expression_bodied_members.rs

using static __Harness;

var n = new Num { V = 3 }
+ new Num { V = 4 }
;
__P((n.Double()).ToString());
__Check("14");

struct Num { public int V; public static Num operator +(Num a, Num b) => new Num { V = a.V + b.V }; public int Double() => V * 2; }

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
