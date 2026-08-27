// vybe-test: csharp/csharp_linq_groupby_join/group_by_clusters_elements_by_key_and_counts_each_group
// origin: languages/csharp/tests/csharp/test_csharp_linq_groupby_join.rs

using static __Harness;

var words = new[] { "apple", "ant", "banana", "bear", "avocado" }
;
var groups = words
    .GroupBy(w => w[0])
    .OrderBy(g => g.Key)
    .Select(g => $"{g.Key}:{g.Count()}");
foreach (var s in groups) __P((s).ToString());
__Check("a:3\nb:2");

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
