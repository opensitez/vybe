// vybe-test: csharp/csharp_comparison_sorting/then_by_breaks_ties_after_primary_sort_key
// origin: languages/csharp/tests/csharp/test_csharp_comparison_sorting.rs

using static __Harness;
using System.Linq;

var values = new[] { "ba", "aa", "c" }
.OrderBy(text => text.Length).ThenBy(text => text);
foreach (var value in values) __P((value).ToString());
__Check("c\naa\nba");

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
