// vybe-test: csharp/csharp_pattern_matching_advanced/null_pattern_matches_missing_nullable_value
// origin: languages/csharp/tests/csharp/test_csharp_pattern_matching_advanced.rs

using static __Harness;

int? value = null;
if (value is null) __P(("missing").ToString());
__Check("missing");

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
