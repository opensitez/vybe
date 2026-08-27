// vybe-test: csharp/csharp_hashset_set_algebra/is_proper_subset_of_true_when_strictly_smaller
// origin: languages/csharp/tests/csharp/test_csharp_hashset_set_algebra.rs

using static __Harness;
using System.Collections.Generic;

var a = new HashSet<int> { 1 }
;
var b = new HashSet<int> { 1, 2 }
;
__P((a.IsProperSubsetOf(b)).ToString());
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
