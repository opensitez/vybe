// vybe-test: csharp/csharp_expression_bodied_members/expr_method_class_returns_bool_comparison
// origin: languages/csharp/tests/csharp/test_csharp_expression_bodied_members.rs

using static __Harness;

__P((new Check().IsZero(0)).ToString());
__P((new Check().IsZero(1)).ToString());
__Check("True\nFalse");

class Check { public bool IsZero(int n) => n == 0; }

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
