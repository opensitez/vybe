// vybe-test: csharp/csharp_switch_expression_core/switch_expr_when_on_type_pattern_int_non_positive
// origin: languages/csharp/tests/csharp/test_csharp_switch_expression_core.rs

using static __Harness;

object o=-1;
__P((o switch{int n when n>0=>"pos",_=>"other"}).ToString());
__Check("other");

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
