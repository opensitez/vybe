// vybe-test: csharp/csharp_hashset_set_algebra/is_subset_of_empty_set_only_for_empty
// origin: languages/csharp/tests/csharp/test_csharp_hashset_set_algebra.rs

using static __Harness;
using System.Collections.Generic;

var empty = new HashSet<int>();
var nonempty = new HashSet<int> { 1 }
;
__P((empty.IsSubsetOf(nonempty)).ToString());
__P((nonempty.IsSubsetOf(empty)).ToString());
__Check("True\nFalse");

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
