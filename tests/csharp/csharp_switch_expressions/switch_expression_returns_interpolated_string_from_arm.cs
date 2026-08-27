// vybe-test: csharp/csharp_switch_expressions/switch_expression_returns_interpolated_string_from_arm
// origin: languages/csharp/tests/csharp/test_csharp_switch_expressions.rs

using static __Harness;

var score = 87;
__P((score switch { >= 90 => $"A:{score}", >= 80 => $"B:{score}", _ => $"C:{score}" }).ToString());
__Check("B:87");

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
