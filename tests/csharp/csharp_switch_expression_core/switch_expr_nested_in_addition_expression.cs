// vybe-test: csharp/csharp_switch_expression_core/switch_expr_nested_in_addition_expression
// origin: languages/csharp/tests/csharp/test_csharp_switch_expression_core.rs

using static __Harness;

int x = 2;
int res = 10 + (x switch { 2 => 20, _ => 0 });
__P(res.ToString());
__Check("30");
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
