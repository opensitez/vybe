// vybe-test: csharp/csharp_hashset_set_algebra/except_then_union_readds_removed_only_if_new
// origin: languages/csharp/tests/csharp/test_csharp_hashset_set_algebra.rs

using static __Harness;
using System.Collections.Generic;

var a = new HashSet<int> { 1, 2, 3 }
;
a.ExceptWith(new[] { 2 });
a.UnionWith(new[] { 2, 4 });
__P((a.Contains(2)).ToString());
__P((a.Contains(4)).ToString());
__Check("True\nTrue");

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
