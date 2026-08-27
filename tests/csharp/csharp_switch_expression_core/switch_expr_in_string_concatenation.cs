// vybe-test: csharp/csharp_switch_expression_core/switch_expr_in_string_concatenation
// origin: languages/csharp/tests/csharp/test_csharp_switch_expression_core.rs

using static __Harness;

var n=4;
__P(("v="+(n switch{4=>"four",_=>"?"})).ToString());
__Check("v=four");

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
