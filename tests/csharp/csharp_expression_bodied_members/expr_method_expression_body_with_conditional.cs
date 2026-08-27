// vybe-test: csharp/csharp_expression_bodied_members/expr_method_expression_body_with_conditional
// origin: languages/csharp/tests/csharp/test_csharp_expression_bodied_members.rs

using static __Harness;

__P((new Sign().Label(-1)).ToString());
__P((new Sign().Label(0)).ToString());
__P((new Sign().Label(2)).ToString());
__Check("neg\nzero\npos");

class Sign { public string Label(int n) => n < 0 ? "neg" : n > 0 ? "pos" : "zero"; }

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
