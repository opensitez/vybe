// vybe-test: csharp/csharp_switch_expressions/switch_expression_matches_string_literal_cases
// origin: languages/csharp/tests/csharp/test_csharp_switch_expressions.rs

using static __Harness;

var word = "beta";
__P((word switch { "alpha" => "A", "beta" => "B", _ => "?" }).ToString());
__Check("B");

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
