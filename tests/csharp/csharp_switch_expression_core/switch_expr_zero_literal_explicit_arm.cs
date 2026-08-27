// vybe-test: csharp/csharp_switch_expression_core/switch_expr_zero_literal_explicit_arm
// origin: languages/csharp/tests/csharp/test_csharp_switch_expression_core.rs

using static __Harness;

var n=0;
__P((n switch{0=>"zero",_=>"nz"}).ToString());
__Check("zero");

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
