// vybe-test: csharp/csharp_switch_expression_core/switch_expr_tuple_pattern_discard_axis
// origin: languages/csharp/tests/csharp/test_csharp_switch_expression_core.rs

using static __Harness;

var p=(3,0);
__P((p switch{(0,0)=>"origin",(_,0)=>"x",_=>"away"}).ToString());
__Check("x");

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
