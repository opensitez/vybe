// vybe-test: csharp/csharp_pattern_matching_advanced/switch_statement_type_pattern_matches_string_arm
// origin: languages/csharp/tests/csharp/test_csharp_pattern_matching_advanced.rs

using static __Harness;

object item = "beta";
switch (item) { case string text: __P((text.ToUpper()).ToString()); break; default: __P(("other").ToString()); break; }
__Check("BETA");

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
