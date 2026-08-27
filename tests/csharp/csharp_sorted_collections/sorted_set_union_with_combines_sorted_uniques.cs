// vybe-test: csharp/csharp_sorted_collections/sorted_set_union_with_combines_sorted_uniques
// origin: languages/csharp/tests/csharp/test_csharp_sorted_collections.rs

using static __Harness;
using System.Collections.Generic;

var a = new SortedSet<int> { 1, 3 }
;
a.UnionWith(new[] { 2, 3, 4 });
__P((a.Count).ToString());
__P((a.Min).ToString());
__P((a.Max).ToString());
__Check("4\n1\n4");

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
