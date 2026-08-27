// vybe-test: csharp/csharp_comparison_sorting/string_comparer_ordinal_ignore_case_treats_text_as_equal
// origin: languages/csharp/tests/csharp/test_csharp_comparison_sorting.rs

using static __Harness;

__P((System.StringComparer.OrdinalIgnoreCase.Equals("abc", "ABC")).ToString());
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
