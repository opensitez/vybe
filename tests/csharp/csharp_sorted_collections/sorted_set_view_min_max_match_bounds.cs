// vybe-test: csharp/csharp_sorted_collections/sorted_set_view_min_max_match_bounds
// origin: languages/csharp/tests/csharp/test_csharp_sorted_collections.rs

using static __Harness;
using System.Collections.Generic;

var ss = new SortedSet<int> { 1, 2, 3, 4, 5, 6 }
;
var view = ss.GetViewBetween(2, 5);
__P((view.Min).ToString());
__P((view.Max).ToString());
__Check("2\n5");

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
