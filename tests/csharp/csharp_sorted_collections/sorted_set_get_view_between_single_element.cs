// vybe-test: csharp/csharp_sorted_collections/sorted_set_get_view_between_single_element
// origin: languages/csharp/tests/csharp/test_csharp_sorted_collections.rs

using static __Harness;
using System.Collections.Generic;

var ss = new SortedSet<int> { 10, 20, 30 }
;
var view = ss.GetViewBetween(20, 20);
__P((view.Min).ToString());
__Check("20");

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
