// vybe-test: csharp/csharp_sorted_collections/sorted_set_reverse_enumeration_not_guaranteed_but_min_max_hold
// origin: languages/csharp/tests/csharp/test_csharp_sorted_collections.rs

using static __Harness;
using System.Collections.Generic;

var ss = new SortedSet<int> { 4, 1, 7 }
;
__P((ss.Min).ToString());
__P((ss.Max).ToString());
__Check("1\n7");

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
