// vybe-test: csharp/csharp_sorted_collections/sorted_set_proper_subset_when_strictly_smaller
// origin: languages/csharp/tests/csharp/test_csharp_sorted_collections.rs

using static __Harness;
using System.Collections.Generic;

var small = new SortedSet<int> { 2 }
;
var big = new SortedSet<int> { 1, 2, 3 }
;
__P((small.IsProperSubsetOf(big)).ToString());
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
