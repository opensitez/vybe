// vybe-test: csharp/csharp_pattern_matching_advanced/constant_pattern_matches_true_boolean
// origin: languages/csharp/tests/csharp/test_csharp_pattern_matching_advanced.rs

using static __Harness;

object value = true;
if (value is true) __P(("yes").ToString());
__Check("yes");

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
