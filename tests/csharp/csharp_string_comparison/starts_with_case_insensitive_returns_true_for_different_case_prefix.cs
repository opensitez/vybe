// vybe-test: csharp/csharp_string_comparison/starts_with_case_insensitive_returns_true_for_different_case_prefix
// origin: languages/csharp/tests/csharp/test_csharp_string_comparison.rs

using static __Harness;

__P(("HELLO".StartsWith("hell",System.StringComparison.OrdinalIgnoreCase)).ToString());
__Check("True");

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
