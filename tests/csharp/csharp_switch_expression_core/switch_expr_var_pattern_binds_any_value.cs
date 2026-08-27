// vybe-test: csharp/csharp_switch_expression_core/switch_expr_var_pattern_binds_any_value
// origin: languages/csharp/tests/csharp/test_csharp_switch_expression_core.rs

using static __Harness;

int x = 42;
string res = x switch {
    var v => $"Val_{v}"
};
__P(res);
__Check("Val_42");
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
