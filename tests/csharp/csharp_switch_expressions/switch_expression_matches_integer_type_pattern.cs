// vybe-test: csharp/csharp_switch_expressions/switch_expression_matches_integer_type_pattern
// origin: languages/csharp/tests/csharp/test_csharp_switch_expressions.rs

using static __Harness;

object item = 7;
__P((item switch { string text => text, int number => (number + 1).ToString(), _ => "other" }).ToString());
__Check("8");

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
