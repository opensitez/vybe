// vybe-test: csharp/csharp_comparison_sorting/array_sort_with_string_comparer_can_ignore_case
// origin: languages/csharp/tests/csharp/test_csharp_comparison_sorting.rs

using static __Harness;

var values = new[] { "b", "A", "c" }
;
System.Array.Sort(values, System.StringComparer.OrdinalIgnoreCase);
foreach (var value in values) __P((value).ToString());
__Check("A\nb\nc");

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
