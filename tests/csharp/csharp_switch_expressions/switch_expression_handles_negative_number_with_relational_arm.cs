// vybe-test: csharp/csharp_switch_expressions/switch_expression_handles_negative_number_with_relational_arm
// origin: languages/csharp/tests/csharp/test_csharp_switch_expressions.rs

using static __Harness;

var x = -3;
__P((x switch { < 0 => "neg", 0 => "zero", > 0 => "pos" }).ToString());
__Check("neg");

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
