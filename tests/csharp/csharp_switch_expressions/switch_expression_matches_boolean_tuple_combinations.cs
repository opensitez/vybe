// vybe-test: csharp/csharp_switch_expressions/switch_expression_matches_boolean_tuple_combinations
// origin: languages/csharp/tests/csharp/test_csharp_switch_expressions.rs

using static __Harness;

var flags = (true, false);
__P((flags switch { (true, true) => "both", (true, false) => "left", _ => "other" }).ToString());
__Check("left");

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
