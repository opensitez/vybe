// vybe-test: csharp/csharp_string_comparison/index_of_with_string_comparison_finds_case_insensitive
// origin: languages/csharp/tests/csharp/test_csharp_string_comparison.rs

using static __Harness;

__P(("fooBAR".IndexOf("bar",System.StringComparison.OrdinalIgnoreCase)).ToString());
__Check("3");

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
