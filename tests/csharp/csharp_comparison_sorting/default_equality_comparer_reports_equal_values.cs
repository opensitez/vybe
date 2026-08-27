// vybe-test: csharp/csharp_comparison_sorting/default_equality_comparer_reports_equal_values
// origin: languages/csharp/tests/csharp/test_csharp_comparison_sorting.rs

using static __Harness;

__P((System.Collections.Generic.EqualityComparer<int>.Default.Equals(4, 4)).ToString());
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
