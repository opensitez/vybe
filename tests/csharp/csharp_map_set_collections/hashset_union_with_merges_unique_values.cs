// vybe-test: csharp/csharp_map_set_collections/hashset_union_with_merges_unique_values
// origin: languages/csharp/tests/csharp/test_csharp_map_set_collections.rs

using static __Harness;
using System.Collections.Generic;

var left = new HashSet<int> { 1, 2 }
;
left.UnionWith(new[] { 2, 3 });
foreach (var item in left) __P((item).ToString());
__Check("1\n2\n3");

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
