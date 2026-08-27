// vybe-test: csharp/csharp_switch_expression_core/switch_expr_nested_switch_as_arm_value
// origin: languages/csharp/tests/csharp/test_csharp_switch_expression_core.rs

using static __Harness;

int x = 1;
string res = x switch {
    1 => (x switch { 1 => "inner", _ => "default" }),
    _ => "outer"
};
__P(res);
__Check("inner");
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
