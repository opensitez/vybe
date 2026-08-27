// vybe-test: csharp/csharp_switch_expressions/switch_expression_uses_parenthesized_result_expression
// origin: languages/csharp/tests/csharp/test_csharp_switch_expressions.rs

using static __Harness;

var x = 4;
__P((x switch { 4 => (2 + 3), _ => 0 }).ToString());
__Check("5");

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
