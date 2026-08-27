// vybe-test: csharp/csharp_comparison_sorting/array_sort_with_parallel_keys_reorders_items_by_key
// origin: languages/csharp/tests/csharp/test_csharp_comparison_sorting.rs

using static __Harness;

var keys = new[] { 2, 1 }
;
var items = new[] { "b", "a" }
;
System.Array.Sort(keys, items);
foreach (var value in items) __P((value).ToString());
__Check("a\nb");

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
