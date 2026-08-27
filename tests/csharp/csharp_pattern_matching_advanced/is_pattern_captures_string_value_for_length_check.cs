// vybe-test: csharp/csharp_pattern_matching_advanced/is_pattern_captures_string_value_for_length_check
// origin: languages/csharp/tests/csharp/test_csharp_pattern_matching_advanced.rs

using static __Harness;

object item = "alpha";
if (item is string text) __P((text.Length).ToString());
__Check("5");

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
