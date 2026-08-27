// vybe-test: csharp/csharp_sorted_collections/sorted_set_overlaps_false_when_disjoint
// origin: languages/csharp/tests/csharp/test_csharp_sorted_collections.rs

using static __Harness;
using System.Collections.Generic;

var a = new SortedSet<int> { 1, 2 }
;
var b = new SortedSet<int> { 5, 6 }
;
__P((a.Overlaps(b)).ToString());
__Check("False");

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
