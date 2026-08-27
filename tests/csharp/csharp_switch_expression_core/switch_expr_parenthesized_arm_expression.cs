// vybe-test: csharp/csharp_switch_expression_core/switch_expr_parenthesized_arm_expression
// origin: languages/csharp/tests/csharp/test_csharp_switch_expression_core.rs

using static __Harness;

var n=2;
__P((n switch{2=>(3+4),_=>0}).ToString());
__Check("7");

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
