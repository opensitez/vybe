// vybe-test: csharp/csharp_switch_expression_core/switch_expr_double_literal_arm
// origin: languages/csharp/tests/csharp/test_csharp_switch_expression_core.rs

using static __Harness;

double d=2.5;
__P((d switch{2.5=>"half",_=>"other"}).ToString());
__Check("half");

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
